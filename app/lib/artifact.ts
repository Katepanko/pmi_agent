import { unzipSync } from "fflate";
import { PDFDocument } from "pdf-lib";
import type { SourceManifestItem } from "./pmi-prompt";
import { conflictSummary, reconcileEvidence, type EvidenceReconciliation } from "./evidence.ts";
import {
  buildPresentationPlanningPrompt,
  parsePresentationModel,
  POWERPOINT_MIME,
  presentationFileName,
  renderPresentation,
  type PresentationModel,
  type PresentationStatus,
} from "./presentation";
import { renderExcelWorkbook } from "./renderers/excel";
import { renderWordDocument } from "./renderers/word";
import { renderPdfReport } from "./renderers/pdf";
import { renderHtmlDashboard } from "./renderers/html";
import type { ArtifactFormat } from "./artifact-intent";
import { DeloitteBrand } from "./branding/deloitte";
import { describeTemplate, type ArtifactTemplate } from "./template.ts";
import { fillCsvTemplate, fillExcelTemplate, fillHtmlTemplate, fillPdfTemplate, fillWordTemplate } from "./report-template.ts";

export { detectArtifactRequest, type ArtifactFormat } from "./artifact-intent";

export type ConsultingReportItem = {
  label: string;
  value?: string;
  detail?: string;
  implication?: string;
  recommendation?: string;
  owner?: string;
  deadline?: string;
  status?: PresentationStatus;
  evidenceType?: "fact" | "calculation" | "inference" | "recommendation" | "gap" | "conflict";
  sourceRefs?: string[];
};

export type ConsultingReportSection = {
  name: string;
  title: string;
  keyMessage?: string;
  type: "summary" | "kpi" | "status" | "milestones" | "risks" | "synergies" | "actions" | "decisions" | "sources" | "detail";
  items: ConsultingReportItem[];
  sourceNotes?: string[];
};

export type ConsultingReportModel = {
  title: string;
  subtitle?: string;
  audience: string;
  reportingPeriod?: string;
  executiveSummary: string;
  sections: ConsultingReportSection[];
};

export type ArtifactContentModel = PresentationModel | ConsultingReportModel;

export type RenderedArtifact = {
  format: ArtifactFormat;
  mimeType: string;
  filename: string;
  bytes: Uint8Array;
  unitCount: number;
  unitLabel: string;
};

export const ARTIFACT_MIME: Record<ArtifactFormat, string> = {
  pptx: POWERPOINT_MIME,
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  html: "text/html; charset=utf-8",
};

function text(value: unknown, fallback = "") {
  return typeof value === "string" ? value.trim() : fallback;
}

function status(value: unknown): PresentationStatus {
  return value === "green" || value === "amber" || value === "red" ? value : "neutral";
}

function evidenceType(value: unknown): ConsultingReportItem["evidenceType"] {
  return value === "fact" || value === "calculation" || value === "inference" || value === "recommendation" || value === "gap" || value === "conflict" ? value : "inference";
}

function parseJsonObject(raw: string) {
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  const candidate = fenced ?? (start >= 0 && end > start ? raw.slice(start, end + 1) : "");
  if (!candidate.trim()) throw new Error("The model did not return a structured artifact model.");
  return JSON.parse(candidate) as Record<string, unknown>;
}

export function parseConsultingReportModel(raw: string, fallbackAudience: string): ConsultingReportModel {
  const value = parseJsonObject(raw);
  const rawSections = Array.isArray(value.sections) ? value.sections : [];
  const sections = rawSections.slice(0, 12).map((rawSection, sectionIndex): ConsultingReportSection => {
    const section = (rawSection && typeof rawSection === "object" ? rawSection : {}) as Record<string, unknown>;
    const rawItems = Array.isArray(section.items) ? section.items : [];
    const items = rawItems.slice(0, 40).map((rawItem): ConsultingReportItem => {
      const item = (rawItem && typeof rawItem === "object" ? rawItem : {}) as Record<string, unknown>;
      return {
        label: text(item.label, "Management point"),
        value: text(item.value) || undefined,
        detail: text(item.detail) || undefined,
        implication: text(item.implication) || undefined,
        recommendation: text(item.recommendation) || undefined,
        owner: text(item.owner) || undefined,
        deadline: text(item.deadline) || undefined,
        status: status(item.status),
        evidenceType: evidenceType(item.evidenceType),
        sourceRefs: Array.isArray(item.sourceRefs) ? item.sourceRefs.filter((ref): ref is string => typeof ref === "string").slice(0, 8) : [],
      };
    });
    const allowedTypes = ["summary", "kpi", "status", "milestones", "risks", "synergies", "actions", "decisions", "sources", "detail"];
    return {
      name: text(section.name, `Section ${sectionIndex + 1}`).slice(0, 40),
      title: text(section.title, `Management update ${sectionIndex + 1}`),
      keyMessage: text(section.keyMessage) || undefined,
      type: allowedTypes.includes(text(section.type)) ? text(section.type) as ConsultingReportSection["type"] : "detail",
      items,
      sourceNotes: Array.isArray(section.sourceNotes) ? section.sourceNotes.filter((note): note is string => typeof note === "string").slice(0, 8) : [],
    };
  });
  if (!sections.length) throw new Error("The artifact model did not contain any report sections.");
  return {
    title: text(value.title, "PMI Management Report"),
    subtitle: text(value.subtitle) || undefined,
    audience: text(value.audience, fallbackAudience || "Management"),
    reportingPeriod: text(value.reportingPeriod) || undefined,
    executiveSummary: text(value.executiveSummary, sections[0].title),
    sections,
  };
}

export function buildConsultingArtifactPrompt(input: {
  request: string;
  format: Exclude<ArtifactFormat, "pptx">;
  audience: string;
  projectContext?: string;
  sources: SourceManifestItem[];
  history: Array<{ role: "user" | "assistant"; content: string }>;
  currentModel?: ConsultingReportModel | null;
  reconciliation?: EvidenceReconciliation;
  template?: ArtifactTemplate | null;
}) {
  const reconciliation = input.reconciliation ?? reconcileEvidence(input.sources);
  const formatPurpose: Record<Exclude<ArtifactFormat, "pptx">, string> = {
    xlsx: "a management workbook with dynamically chosen sheets, a first-sheet executive view, auditable detail, filters, frozen headers, semantic status formatting, and formulas only where evidence supports them",
    docx: "a skimmable executive document with a professional title block, conclusion-led headings, management callouts, real tables where comparison is useful, headers, footers, page numbers, and source notes",
    pdf: "a finished management report with deliberate page composition, executive summary, concise callouts, well-spaced tables, status indicators, page numbers, and source notes",
    html: "a self-contained responsive executive dashboard with strong hierarchy, concise KPI/status views, insight-led sections, risks, decisions, and an evidence register",
  };
  const manifest = input.sources.map((source) => ({
    id: source.id,
    file_name: source.fileName,
    extraction_status: source.status,
    available_evidence: source.excerpt ?? "No extracted excerpt supplied.",
    warnings: source.warnings ?? [],
  }));
  return `You are the shared consulting-report engine for a senior post-merger integration team.

Design the content model for ${formatPurpose[input.format]}. The renderer will create the actual file; do not return planning prose or Markdown. Use Minto / Pyramid Principle logic: lead with the governing thought, use conclusion-oriented section titles, transform data into findings and implications, and make decisions/actions explicit. Adapt content density to ${input.audience || "the inferred audience"}.

Evidence discipline:
- Never invent numerical values, dates, owners, status, targets, or KPIs.
- Distinguish facts, reproducible calculations, consultant inferences, recommendations, and gaps with evidenceType.
- Use sourceRefs with supplied source IDs for factual items.
- State "Not evidenced" for material missing information.
- Preserve extraction warnings, conflicts, and source limitations.
- Treat the deterministic reconciliation below as mandatory. Do not select or average an unresolved conflicting value. Include every material conflict naturally in the relevant management section with values and provenance.
- Recommendations must be supportable and require validation.
- Keep item text concise; avoid walls of text and raw dumps.

Template discipline:
- ${input.template ? `Use the explicitly selected ${input.template.fileType.toUpperCase()} file as the report template. Match its section order, field/table semantics, density, and visual hierarchy where the requested output format supports them.` : "No user template was selected. Use the application's standard report system."}
- The template is a structural/style reference, never PMI deal evidence. Do not carry over example companies, figures, dates, owners, statuses, or conclusions.
- When a semantic layout model is supplied, plan records for its named fields and table headers, and keep each section within the stated row/item capacity. Prefer an explicit evidence gap to unsupported filler.

Return ONLY one valid JSON object with this shape:
{
  "title": "deliverable title",
  "subtitle": "optional context",
  "audience": "audience",
  "reportingPeriod": "optional source-backed period",
  "executiveSummary": "one-sentence governing thought",
  "sections": [{
    "name": "short format-native section or sheet name",
    "title": "conclusion-oriented management heading",
    "keyMessage": "optional concise implication",
    "type": "summary|kpi|status|milestones|risks|synergies|actions|decisions|sources|detail",
    "items": [{
      "label": "short heading or record name",
      "value": "optional source-backed value",
      "detail": "source-backed evidence",
      "implication": "what it means for management",
      "recommendation": "optional action; validation required",
      "owner": "optional source-backed owner",
      "deadline": "optional source-backed deadline",
      "status": "green|amber|red|neutral",
      "evidenceType": "fact|calculation|inference|recommendation|gap|conflict",
      "sourceRefs": ["source-id"]
    }],
    "sourceNotes": ["short evidence limitation or source note"]
  }]
}

${input.currentModel ? "This is a revision. Preserve unaffected sections and make only the requested changes. Return the complete revised model." : "This is a new artifact. Choose the smallest set of sections that fully answers the management question."}

Format: ${input.format}
Audience: ${input.audience || "Infer from request"}
Project context: ${input.projectContext || "No project context supplied."}
Complete source manifest: ${JSON.stringify(manifest)}
Selected template (excluded from the evidence manifest): ${input.template ? JSON.stringify(describeTemplate(input.template)) : "None"}
Deterministic cross-source reconciliation: ${JSON.stringify(conflictSummary(reconciliation))}
Prior conversation: ${JSON.stringify(input.history.slice(-16))}
Current artifact model: ${input.currentModel ? JSON.stringify(input.currentModel) : "None"}
User request: ${input.request}`;
}

function cleanFilePart(value: string) {
  return value.replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 64) || "PMI_Report";
}

export function artifactFileName(title: string, format: ArtifactFormat, version: number) {
  return `${cleanFilePart(title)}${version > 1 ? `_v${version}` : ""}.${format}`;
}

function toUint8Array(value: Uint8Array | ArrayBuffer | Buffer) {
  if (value instanceof Uint8Array) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  return new Uint8Array(value);
}

export async function renderArtifact(input: {
  format: ArtifactFormat;
  model: ArtifactContentModel;
  version: number;
  sources: SourceManifestItem[];
  template?: ArtifactTemplate | null;
}): Promise<RenderedArtifact> {
  if (input.format === "pptx") {
    const model = input.model as PresentationModel;
    const bytes = await renderPresentation(model, input.template);
    const renderedSlideCount = Object.keys(unzipSync(bytes)).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).length;
    return { format: "pptx", mimeType: ARTIFACT_MIME.pptx, filename: presentationFileName(model, input.version), bytes, unitCount: renderedSlideCount, unitLabel: "slides" };
  }
  const model = input.model as ConsultingReportModel;
  const filename = artifactFileName(model.title, input.format, input.version);
  if (input.format === "xlsx") {
    if (input.template?.fileType === "xlsx" && input.template.bytes) {
      const bytes = await fillExcelTemplate(input.template.bytes, model);
      return { format: "xlsx", mimeType: ARTIFACT_MIME.xlsx, filename, bytes, unitCount: model.sections.length, unitLabel: "sections" };
    }
    if (input.template?.fileType === "csv" && input.template.bytes) {
      const bytes = await fillCsvTemplate(input.template.bytes, model);
      return { format: "xlsx", mimeType: ARTIFACT_MIME.xlsx, filename, bytes, unitCount: model.sections.length, unitLabel: "sections" };
    }
    const result = await renderExcelWorkbook(model, input.sources);
    return { format: "xlsx", mimeType: ARTIFACT_MIME.xlsx, filename, bytes: toUint8Array(result.bytes), unitCount: result.sheetCount, unitLabel: "sheets" };
  }
  if (input.format === "docx") {
    if (input.template?.fileType === "docx" && input.template.bytes) {
      const bytes = fillWordTemplate(input.template.bytes, model);
      return { format: "docx", mimeType: ARTIFACT_MIME.docx, filename, bytes, unitCount: model.sections.length, unitLabel: "sections" };
    }
    const result = await renderWordDocument(model, input.sources);
    return { format: "docx", mimeType: ARTIFACT_MIME.docx, filename, bytes: toUint8Array(result.bytes), unitCount: result.sectionCount, unitLabel: "sections" };
  }
  if (input.format === "pdf") {
    if (input.template?.fileType === "pdf" && input.template.bytes) {
      const result = await fillPdfTemplate(input.template.bytes, model);
      return { format: "pdf", mimeType: ARTIFACT_MIME.pdf, filename, bytes: result.bytes, unitCount: result.pageCount, unitLabel: "pages" };
    }
    const result = await renderPdfReport(model, input.sources);
    return { format: "pdf", mimeType: ARTIFACT_MIME.pdf, filename, bytes: result.bytes, unitCount: result.pageCount, unitLabel: "pages" };
  }
  const html = input.template && ["html", "htm"].includes(input.template.fileType) && input.template.bytes
    ? fillHtmlTemplate(input.template.bytes, model)
    : renderHtmlDashboard(model, input.sources);
  return { format: "html", mimeType: ARTIFACT_MIME.html, filename, bytes: new TextEncoder().encode(html), unitCount: model.sections.length, unitLabel: "sections" };
}

export async function validateArtifact(artifact: RenderedArtifact, template?: ArtifactTemplate | null) {
  if (!artifact.bytes.byteLength) throw new Error(`${artifact.format.toUpperCase()} rendering produced an empty file.`);
  if (!artifact.filename.toLowerCase().endsWith(`.${artifact.format}`)) throw new Error("Artifact filename does not match its requested format.");
  if (artifact.format === "pdf") {
    const document = await PDFDocument.load(artifact.bytes);
    if (document.getPageCount() < 1) throw new Error("The generated PDF has no pages.");
    const pdfText = new TextDecoder("latin1").decode(artifact.bytes);
    if (!template && (!/\/Subtype\s*\/Image/.test(pdfText) || document.getAuthor() !== DeloitteBrand.name)) throw new Error("The generated PDF is missing Deloitte branding.");
    return;
  }
  if (artifact.format === "html") {
    const html = new TextDecoder().decode(artifact.bytes);
    if (!/<!doctype html>/i.test(html) || !/<html[\s>]/i.test(html) || !/<\/html>/i.test(html)) throw new Error("The generated HTML document is invalid.");
    if (!template && (!/<img[^>]+alt="Deloitte"/i.test(html) || !/<img[^>]+src="data:image\/png;base64,/i.test(html))) throw new Error("The generated HTML is missing the embedded Deloitte logo.");
    return;
  }
  const files = unzipSync(artifact.bytes);
  const required = artifact.format === "pptx"
    ? ["[Content_Types].xml", "ppt/presentation.xml", "ppt/slides/slide1.xml"]
    : artifact.format === "xlsx"
      ? ["[Content_Types].xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml"]
      : ["[Content_Types].xml", "word/document.xml"];
  for (const path of required) if (!files[path]) throw new Error(`The generated ${artifact.format.toUpperCase()} package is missing ${path}.`);
  if (artifact.format === "docx" && !template) {
    const header = Object.entries(files).filter(([path]) => /^word\/header\d+\.xml$/.test(path)).map(([, bytes]) => new TextDecoder().decode(bytes)).join("\n");
    if (!/<w:drawing>/i.test(header) || !Object.keys(files).some((path) => /^word\/media\/.+\.png$/i.test(path))) throw new Error("The generated Word document is missing the Deloitte header logo.");
  }
  if (artifact.format === "xlsx" && !template) {
    if (!Object.keys(files).some((path) => /^xl\/media\/.+\.png$/i.test(path)) || !files["xl/drawings/drawing1.xml"]) throw new Error("The generated Excel workbook is missing Deloitte branding on the executive sheet.");
  }
}

export function planArtifact(input: {
  format: ArtifactFormat;
  request: string;
  audience: string;
  projectContext?: string;
  sources: SourceManifestItem[];
  history: Array<{ role: "user" | "assistant"; content: string }>;
  currentModel?: ArtifactContentModel | null;
  reconciliation?: EvidenceReconciliation;
  template?: ArtifactTemplate | null;
}) {
  if (input.format === "pptx") {
    return buildPresentationPlanningPrompt({ ...input, currentPresentation: input.currentModel as PresentationModel | null | undefined });
  }
  return buildConsultingArtifactPrompt({ ...input, format: input.format, currentModel: input.currentModel as ConsultingReportModel | null | undefined });
}

export function parseArtifactModel(format: ArtifactFormat, raw: string, audience: string): ArtifactContentModel {
  return format === "pptx" ? parsePresentationModel(raw, audience) : parseConsultingReportModel(raw, audience);
}

export function artifactTitle(model: ArtifactContentModel) {
  return model.title;
}
