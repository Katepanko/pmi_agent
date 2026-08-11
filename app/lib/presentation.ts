import PptxGenJS from "pptxgenjs";
import type { SourceManifestItem } from "./pmi-prompt";
import { applyDeloittePowerPointTemplate, DeloitteBrand, validateDeloittePowerPoint } from "./branding/deloitte.ts";
import { conflictSummary, reconcileEvidence, type EvidenceReconciliation } from "./evidence.ts";

export const POWERPOINT_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export type PresentationStatus = "green" | "amber" | "red" | "neutral";

export type PresentationItem = {
  label: string;
  value?: string;
  detail?: string;
  implication?: string;
  owner?: string;
  deadline?: string;
  status?: PresentationStatus;
  evidenceType?: "fact" | "calculation" | "inference" | "recommendation" | "gap" | "conflict";
  sourceRefs?: string[];
};

export type PresentationSlide = {
  title: string;
  kicker?: string;
  keyMessage?: string;
  layout?: "cover" | "summary" | "trajectory" | "risks" | "synergies" | "decisions" | "timeline" | "comparison" | "cards";
  items: PresentationItem[];
  sourceNotes?: string[];
};

export type PresentationModel = {
  title: string;
  subtitle?: string;
  projectName?: string;
  location?: string;
  date?: string;
  audience: string;
  executiveSummary: string;
  slides: PresentationSlide[];
};

const STATUS_COLORS: Record<PresentationStatus, string> = {
  green: DeloitteBrand.colors.green,
  amber: DeloitteBrand.colors.amber,
  red: DeloitteBrand.colors.red,
  neutral: DeloitteBrand.colors.coolGray,
};

const COLORS = {
  ink: DeloitteBrand.colors.black,
  body: "313131",
  muted: DeloitteBrand.colors.coolGray,
  green: DeloitteBrand.colors.brightGreen,
  greenDark: DeloitteBrand.colors.deepGreen,
  greenPale: DeloitteBrand.colors.paleGreen,
  paper: "F7F7F7",
  white: DeloitteBrand.colors.white,
  line: DeloitteBrand.colors.lightGray,
  amberPale: "FFF4DE",
  redPale: "FDECEC",
};

const REQUEST_ACTION = /\b(create|generate|prepare|make|build|produce|turn|draft|develop|put together)\b/i;
const PRESENTATION_NOUN = /\b(power\s*point|pptx?|presentation|slide\s*deck|deck|slides?)\b/i;
const REVISION_ACTION = /\b(change|revise|update|edit|adjust|refine|replace|remove|add|rework|make)\b/i;

export function isPresentationRequest(message: string, hasPriorPresentation = false) {
  const explicitRequest = PRESENTATION_NOUN.test(message) && (REQUEST_ACTION.test(message) || /\b\d+\s+slides?\b/i.test(message));
  const revisionRequest = hasPriorPresentation && REVISION_ACTION.test(message) && PRESENTATION_NOUN.test(message);
  return explicitRequest || revisionRequest;
}

function desiredSlideCount(message: string) {
  const match = message.match(/\b(?:create|prepare|make|generate|build|produce)?\s*(\d{1,2})\s+slides?\b/i);
  if (!match) return null;
  const count = Number(match[1]);
  return count >= 1 && count <= 15 ? count : null;
}

export function buildPresentationPlanningPrompt(input: {
  request: string;
  audience: string;
  projectContext?: string;
  sources: SourceManifestItem[];
  history: Array<{ role: "user" | "assistant"; content: string }>;
  currentPresentation?: PresentationModel | null;
  reconciliation?: EvidenceReconciliation;
}) {
  const reconciliation = input.reconciliation ?? reconcileEvidence(input.sources);
  const requestedCount = desiredSlideCount(input.request);
  const manifest = input.sources.map((source) => ({
    id: source.id,
    file_name: source.fileName,
    extraction_status: source.status,
    available_evidence: source.excerpt ?? "No extracted excerpt supplied.",
    warnings: source.warnings ?? [],
  }));
  return `You are the storyline planner for a senior post-merger integration consulting team.

Create a decision-ready, editable presentation model. Use Minto / Pyramid Principle logic: each slide title must state its conclusion, supporting items must be MECE where useful, and implications must explain why the evidence matters. Adapt the storyline, slide count, ordering, and visuals to the request and evidence; never use a fixed slide sequence.

Evidence discipline:
- Never invent numerical values, dates, owners, statuses, or targets.
- Distinguish source facts, calculations, consultant inferences, recommendations, and gaps using evidenceType.
- Use sourceRefs with supplied source IDs for factual items.
- Label recommendations as requiring validation.
- Surface material uncertainty and incomplete extraction.
- Treat the deterministic reconciliation below as mandatory. Never select or average an unresolved conflicting value; show every material conflict with all source values and provenance.
- Keep text concise enough for a management slide: no item detail over 42 words, no slide title over 18 words, and no more than 6 items per slide.

Return ONLY one valid JSON object with this exact shape:
{
  "title": "file/deck title",
  "subtitle": "optional reporting period or context",
  "projectName": "optional source-backed project or client name",
  "location": "optional source-backed location",
  "date": "optional source-backed presentation date",
  "audience": "audience",
  "executiveSummary": "one-sentence governing thought",
  "slides": [
    {
      "title": "action title stating a conclusion",
      "kicker": "optional section label",
      "keyMessage": "optional short implication",
      "layout": "cover|summary|trajectory|risks|synergies|decisions|timeline|comparison|cards",
      "items": [{
        "label": "short heading",
        "value": "optional concise value",
        "detail": "source-backed evidence",
        "implication": "optional management implication",
        "owner": "optional source-backed owner",
        "deadline": "optional source-backed deadline",
        "status": "green|amber|red|neutral",
        "evidenceType": "fact|calculation|inference|recommendation|gap|conflict",
        "sourceRefs": ["source-id"]
      }],
      "sourceNotes": ["short source or uncertainty note"]
    }
  ]
}

${requestedCount ? `The user requested ${requestedCount} total slides. Return exactly ${requestedCount} slide objects; the renderer will not add a separate cover slide.` : "Choose the smallest slide count that fully answers the request, normally 4–7 slides."}
${input.currentPresentation ? "This is a revision. Preserve every unaffected slide and change only what the user requested. The new output must remain a complete presentation." : "This is a new presentation."}

Audience: ${input.audience || "Infer from request"}
Project context: ${input.projectContext || "No project context supplied."}
Complete source manifest: ${JSON.stringify(manifest)}
Deterministic cross-source reconciliation: ${JSON.stringify(conflictSummary(reconciliation))}
Prior conversation: ${JSON.stringify(input.history.slice(-16))}
Current presentation: ${input.currentPresentation ? JSON.stringify(input.currentPresentation) : "None"}
User request: ${input.request}`;
}

function text(value: unknown, fallback = "") {
  return typeof value === "string" ? value.trim() : fallback;
}

function status(value: unknown): PresentationStatus {
  return value === "green" || value === "amber" || value === "red" ? value : "neutral";
}

function evidenceType(value: unknown): PresentationItem["evidenceType"] {
  return value === "fact" || value === "calculation" || value === "inference" || value === "recommendation" || value === "gap" || value === "conflict" ? value : "inference";
}

function parseJsonObject(raw: string) {
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
  const candidate = fenced ?? raw.slice(raw.indexOf("{"), raw.lastIndexOf("}") + 1);
  if (!candidate.trim()) throw new Error("The model did not return a presentation model.");
  return JSON.parse(candidate) as Record<string, unknown>;
}

export function parsePresentationModel(raw: string, fallbackAudience: string): PresentationModel {
  const value = parseJsonObject(raw);
  const rawSlides = Array.isArray(value.slides) ? value.slides : [];
  const slides = rawSlides.slice(0, 15).map((rawSlide, slideIndex): PresentationSlide => {
    const slide = (rawSlide && typeof rawSlide === "object" ? rawSlide : {}) as Record<string, unknown>;
    const rawItems = Array.isArray(slide.items) ? slide.items : [];
    const items = rawItems.slice(0, 6).map((rawItem): PresentationItem => {
      const item = (rawItem && typeof rawItem === "object" ? rawItem : {}) as Record<string, unknown>;
      return {
        label: text(item.label, "Management point"),
        value: text(item.value) || undefined,
        detail: text(item.detail) || undefined,
        implication: text(item.implication) || undefined,
        owner: text(item.owner) || undefined,
        deadline: text(item.deadline) || undefined,
        status: status(item.status),
        evidenceType: evidenceType(item.evidenceType),
        sourceRefs: Array.isArray(item.sourceRefs) ? item.sourceRefs.filter((ref): ref is string => typeof ref === "string").slice(0, 5) : [],
      };
    });
    return {
      title: text(slide.title, `Management update ${slideIndex + 1}`),
      kicker: text(slide.kicker) || undefined,
      keyMessage: text(slide.keyMessage) || undefined,
      layout: ["cover", "summary", "trajectory", "risks", "synergies", "decisions", "timeline", "comparison", "cards"].includes(text(slide.layout))
        ? text(slide.layout) as PresentationSlide["layout"]
        : "cards",
      items,
      sourceNotes: Array.isArray(slide.sourceNotes) ? slide.sourceNotes.filter((note): note is string => typeof note === "string").slice(0, 4) : [],
    };
  });
  if (!slides.length) throw new Error("The presentation model did not contain any slides.");
  return {
    title: text(value.title, "PMI Management Report"),
    subtitle: text(value.subtitle) || undefined,
    projectName: text(value.projectName) || undefined,
    location: text(value.location) || undefined,
    date: text(value.date) || undefined,
    audience: text(value.audience, fallbackAudience || "Management"),
    executiveSummary: text(value.executiveSummary, slides[0].title),
    slides,
  };
}

function cleanFilePart(value: string) {
  return value.replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 64) || "PMI_Report";
}

export function presentationFileName(model: PresentationModel, version: number) {
  return `${cleanFilePart(model.title)}${version > 1 ? `_v${version}` : ""}.pptx`;
}

function addSlideHeader(slide: PptxGenJS.Slide, current: PresentationSlide) {
  const titleLength = current.title.trim().length;
  const titleFontSize = titleLength <= 60 ? 30 : titleLength <= 100 ? 27 : 24;
  const titleHeight = titleLength <= 60 ? 0.58 : 0.88;
  const titleX = 0.55;
  const titleY = 0.58;
  const subtitleY = titleY + titleHeight + 0.06;

  slide.addText((current.kicker ?? "MANAGEMENT UPDATE").toUpperCase(), { x: titleX, y: 0.21, w: 11.7, h: 0.2, fontFace: DeloitteBrand.typography.body, fontSize: 8.5, bold: true, color: COLORS.greenDark, charSpacing: 1.6, margin: 0 });
  slide.addText(current.title, { x: titleX, y: titleY, w: 12.05, h: titleHeight, fontFace: "Aptos Display", fontSize: titleFontSize, bold: true, color: COLORS.ink, breakLine: false, margin: 0, valign: "top", fit: "shrink" });
  if (current.keyMessage) {
    slide.addText(current.keyMessage, { x: titleX, y: subtitleY, w: 11.7, h: 0.34, fontFace: "Aptos", fontSize: 11.5, color: COLORS.body, margin: 0, valign: "top", fit: "shrink" });
  }
}

function addEvidenceLabel(slide: PptxGenJS.Slide, item: PresentationItem, x: number, y: number) {
  const label = item.evidenceType === "recommendation" ? "AI RECOMMENDATION · VALIDATE" : (item.evidenceType ?? "inference").toUpperCase();
  slide.addText(label, { x, y, w: 2.15, h: 0.16, fontFace: "Aptos", fontSize: 6.5, bold: true, color: item.evidenceType === "recommendation" ? COLORS.greenDark : COLORS.muted, charSpacing: 0.5, margin: 0, fit: "shrink" });
}

function addCards(slide: PptxGenJS.Slide, current: PresentationSlide) {
  const items = current.items.length ? current.items : [{ label: "Evidence gap", detail: "No supported detail was available for this page.", evidenceType: "gap" as const, status: "neutral" as const }];
  const columns = items.length <= 3 ? items.length : 3;
  const rows = Math.ceil(items.length / columns);
  const gap = 0.22;
  const width = (12 - gap * (columns - 1)) / columns;
  const height = Math.min(2.08, (4.72 - gap * (rows - 1)) / rows);
  items.forEach((item, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = 0.66 + column * (width + gap);
    const y = 2.05 + row * (height + gap);
    const accent = STATUS_COLORS[item.status ?? "neutral"];
    slide.addShape("rect", { x, y, w: width, h: height, rectRadius: 0.04, line: { color: COLORS.line, width: 0.7 }, fill: { color: COLORS.white } });
    slide.addShape("rect", { x, y, w: 0.07, h: height, line: { color: accent, transparency: 100 }, fill: { color: accent } });
    if (item.value) slide.addText(item.value, { x: x + 0.24, y: y + 0.22, w: width - 0.46, h: 0.38, fontFace: "Aptos Display", fontSize: 20, bold: true, color: accent, margin: 0, fit: "shrink" });
    slide.addText(item.label, { x: x + 0.24, y: y + (item.value ? 0.64 : 0.3), w: width - 0.46, h: 0.42, fontFace: "Aptos", fontSize: 11.5, bold: true, color: COLORS.ink, margin: 0, fit: "shrink" });
    const detail = [item.detail, item.implication ? `Implication: ${item.implication}` : ""].filter(Boolean).join("\n");
    slide.addText(detail || "Evidence to be confirmed.", { x: x + 0.24, y: y + (item.value ? 1.05 : 0.78), w: width - 0.46, h: height - (item.value ? 1.34 : 1.05), fontFace: "Aptos", fontSize: 8.8, color: COLORS.body, breakLine: false, margin: 0, valign: "top", fit: "shrink", paraSpaceAfterPt: 4 });
    addEvidenceLabel(slide, item, x + 0.24, y + height - 0.25);
  });
}

function addManagementTable(slide: PptxGenJS.Slide, current: PresentationSlide) {
  const decisions = current.layout === "decisions";
  const headers = decisions ? ["DECISION / ESCALATION", "WHY IT MATTERS", "OWNER", "BY WHEN"] : ["RISK / ISSUE", "BUSINESS IMPLICATION", "MITIGATION / OWNER", "STATUS"];
  const widths = decisions ? [4.15, 4.35, 1.65, 1.45] : [3.35, 4.2, 3.15, 0.9];
  let x = 0.66;
  headers.forEach((header, index) => {
    slide.addShape("rect", { x, y: 2.03, w: widths[index], h: 0.42, line: { color: COLORS.ink, transparency: 100 }, fill: { color: COLORS.ink } });
    slide.addText(header, { x: x + 0.11, y: 2.15, w: widths[index] - 0.2, h: 0.13, fontFace: "Aptos", fontSize: 7.2, bold: true, color: COLORS.white, charSpacing: 0.45, margin: 0, fit: "shrink" });
    x += widths[index];
  });
  const items = current.items.slice(0, 5);
  const rowHeight = Math.min(0.84, 4.26 / Math.max(items.length, 1));
  items.forEach((item, row) => {
    const y = 2.45 + row * rowHeight;
    const fill = row % 2 === 0 ? COLORS.white : COLORS.paper;
    slide.addShape("rect", { x: 0.66, y, w: 11.6, h: rowHeight, line: { color: COLORS.line, width: 0.5 }, fill: { color: fill } });
    const cells = decisions
      ? [item.label, item.implication ?? item.detail ?? "Validation required", item.owner ?? "TBC", item.deadline ?? "TBC"]
      : [item.label, item.implication ?? item.detail ?? "Validation required", [item.detail, item.owner ? `Owner: ${item.owner}` : ""].filter(Boolean).join(" · "), (item.status ?? "neutral").toUpperCase()];
    let cellX = 0.66;
    cells.forEach((cell, index) => {
      slide.addText(cell, { x: cellX + 0.1, y: y + 0.11, w: widths[index] - 0.2, h: rowHeight - 0.2, fontFace: "Aptos", fontSize: index === 0 ? 9.4 : 8.2, bold: index === 0, color: index === 3 && !decisions ? STATUS_COLORS[item.status ?? "neutral"] : COLORS.body, align: index === 3 && !decisions ? "center" : "left", margin: 0, valign: "mid", fit: "shrink" });
      cellX += widths[index];
    });
  });
}

function addTrajectory(slide: PptxGenJS.Slide, current: PresentationSlide) {
  const items = current.items.slice(0, 6);
  const rowHeight = Math.min(0.78, 4.35 / Math.max(items.length, 1));
  items.forEach((item, index) => {
    const y = 2.08 + index * rowHeight;
    const accent = STATUS_COLORS[item.status ?? "neutral"];
    slide.addText(item.label, { x: 0.72, y: y + 0.1, w: 2.55, h: rowHeight - 0.18, fontFace: "Aptos", fontSize: 10, bold: true, color: COLORS.ink, margin: 0, valign: "mid", fit: "shrink" });
    slide.addShape("rect", { x: 3.38, y: y + 0.22, w: 2.9, h: 0.2, rectRadius: 0.06, line: { color: COLORS.line, transparency: 100 }, fill: { color: COLORS.line } });
    const progressWidth = item.status === "green" ? 2.55 : item.status === "amber" ? 1.75 : item.status === "red" ? 1.05 : 1.45;
    slide.addShape("rect", { x: 3.38, y: y + 0.22, w: progressWidth, h: 0.2, rectRadius: 0.06, line: { color: accent, transparency: 100 }, fill: { color: accent } });
    slide.addText([item.detail, item.implication].filter(Boolean).join(" · ") || "Status evidence to be confirmed", { x: 6.55, y: y + 0.08, w: 5.46, h: rowHeight - 0.15, fontFace: "Aptos", fontSize: 8.4, color: COLORS.body, margin: 0, valign: "mid", fit: "shrink" });
    slide.addShape("line", { x: 0.72, y: y + rowHeight - 0.03, w: 11.28, h: 0, line: { color: COLORS.line, width: 0.5 } });
  });
}

function addSourceNotes(slide: PptxGenJS.Slide, notes: string[] | undefined) {
  if (!notes?.length) return;
  slide.addText(`Sources / caveats: ${notes.join("  ·  ")}`, { x: 0.66, y: 6.78, w: 11.35, h: 0.2, fontFace: "Aptos", fontSize: 6.6, italic: true, color: COLORS.muted, margin: 0, fit: "shrink" });
}

function addCover(slide: PptxGenJS.Slide, model: PresentationModel, current: PresentationSlide) {
  slide.background = { color: COLORS.ink };
  slide.addText((current.kicker ?? model.audience).toUpperCase(), { x: 0.82, y: 0.75, w: 10.8, h: 0.25, fontFace: "Aptos", fontSize: 9.5, bold: true, color: "82D9AA", charSpacing: 1.8, margin: 0 });
  slide.addText(current.title || model.title, { x: 0.82, y: 1.35, w: 10.9, h: 1.75, fontFace: "Aptos Display", fontSize: 32, bold: true, color: COLORS.white, margin: 0, valign: "mid", fit: "shrink" });
  slide.addText(current.keyMessage ?? model.executiveSummary, { x: 0.84, y: 3.35, w: 9.55, h: 0.9, fontFace: "Aptos", fontSize: 15, color: "DDE6E0", margin: 0, fit: "shrink" });
  slide.addShape("line", { x: 0.84, y: 5.65, w: 2.25, h: 0, line: { color: COLORS.green, width: 3 } });
  slide.addText(model.subtitle ?? "Decision-ready integration reporting", { x: 0.84, y: 5.82, w: 7.8, h: 0.3, fontFace: "Aptos", fontSize: 10, color: "AFC0B5", margin: 0 });
  slide.addText(`PMI Agent  ·  ${model.audience}`, { x: 0.84, y: 6.92, w: 5.5, h: 0.18, fontFace: "Aptos", fontSize: 7.5, color: "8FA096", charSpacing: 0.4, margin: 0 });
}

export async function renderPresentation(model: PresentationModel) {
  const slides = model.slides[0]?.layout === "cover" ? model.slides : [{
    title: model.title,
    keyMessage: model.executiveSummary,
    layout: "cover" as const,
    items: [],
    sourceNotes: [],
  }, ...model.slides];
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = DeloitteBrand.name;
  pptx.company = DeloitteBrand.name;
  pptx.subject = model.executiveSummary;
  pptx.title = model.title;
  pptx.lang = "en-US";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "en-US",
  };
  pptx.defineSlideMaster({
    title: "PMI_CONSULTING",
    background: { color: COLORS.white },
    objects: [],
  });

  slides.forEach((current) => {
    const slide = pptx.addSlide({ masterName: "PMI_CONSULTING" });
    if (current.layout === "cover") {
      addCover(slide, model, current);
      slide.addNotes(`[Sources]\n${current.sourceNotes?.length ? current.sourceNotes.map((note) => `- ${note}`).join("\n") : "- No external claim on this cover slide."}`);
      return;
    }
    addSlideHeader(slide, current);
    if (current.layout === "risks" || current.layout === "decisions") addManagementTable(slide, current);
    else if (current.layout === "trajectory" || current.layout === "timeline") addTrajectory(slide, current);
    else addCards(slide, current);
    addSourceNotes(slide, current.sourceNotes);
    const sourceIds = [...new Set(current.items.flatMap((item) => item.sourceRefs ?? []))];
    const notes = [...sourceIds.map((id) => `Source ID: ${id}`), ...(current.sourceNotes ?? [])];
    slide.addNotes(`[Sources]\n${notes.length ? notes.map((note) => `- ${note}`).join("\n") : "- No source reference supplied; validate claims before circulation."}`);
  });

  const result = await pptx.write({ outputType: "uint8array", compression: true });
  const generated = result instanceof Uint8Array
    ? result
    : result instanceof ArrayBuffer
      ? new Uint8Array(result)
      : result instanceof Blob
        ? new Uint8Array(await result.arrayBuffer())
        : null;
  if (!generated) throw new Error("The PowerPoint renderer returned an unsupported output type.");
  const branded = applyDeloittePowerPointTemplate({
    generated,
    title: model.title,
    subtitle: model.subtitle,
    projectName: model.projectName,
    location: model.location,
    date: model.date,
    audience: model.audience,
    slideCount: slides.length,
  });
  validateDeloittePowerPoint(branded, {
    title: model.title,
    slideCount: slides.length,
    metadata: [model.projectName, model.subtitle, model.audience, model.location, model.date].filter((value): value is string => Boolean(value)),
  });
  return branded;
}
