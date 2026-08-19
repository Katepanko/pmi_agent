import type { SourceManifestItem } from "./pmi-prompt";
import type { ArtifactFormat } from "./artifact-intent";

export const TEMPLATE_FILE_TYPES = new Set(["pptx", "xlsx", "xls", "csv", "docx", "pdf", "html", "htm", "png", "jpg", "jpeg"]);

export type ArtifactTemplate = {
  sourceId: string;
  fileName: string;
  fileType: string;
  status: SourceManifestItem["status"];
  excerpt?: string;
  metadata?: Record<string, unknown>;
  warnings?: string[];
  bytes?: Uint8Array;
  layoutModel?: Record<string, unknown>;
};

function normalized(value: string) {
  return value.normalize("NFKC").toLocaleLowerCase();
}

export function messageMentionsTemplate(message: string, fileName: string) {
  const haystack = normalized(message);
  const fullName = normalized(fileName);
  const baseName = fullName.replace(/\.[^.]+$/, "");
  return [fullName, baseName].some((candidate) => {
    const marker = `@${candidate}`;
    let start = haystack.indexOf(marker);
    while (start >= 0) {
      const following = haystack[start + marker.length];
      if (!following || !/[\p{L}\p{N}_-]/u.test(following)) return true;
      start = haystack.indexOf(marker, start + marker.length);
    }
    return false;
  });
}

/** Resolve the single attachment explicitly marked with @filename as a template. */
export function resolveTemplateReference(message: string, sources: SourceManifestItem[]): ArtifactTemplate | null {
  const matches = sources
    .filter((source) => TEMPLATE_FILE_TYPES.has((source.fileType ?? source.fileName.split(".").pop() ?? "").toLocaleLowerCase()))
    .filter((source) => messageMentionsTemplate(message, source.fileName));

  if (matches.length > 1) {
    throw new Error(`Choose one template per request. Multiple @file references matched: ${matches.map((source) => source.fileName).join(", ")}.`);
  }
  const source = matches[0];
  if (!source) return null;
  return {
    sourceId: source.id,
    fileName: source.fileName,
    fileType: (source.fileType ?? source.fileName.split(".").pop() ?? "").toLocaleLowerCase(),
    status: source.status,
    excerpt: source.excerpt,
    metadata: source.metadata,
    warnings: source.warnings,
  };
}

export function hasUnresolvedTemplateDirective(message: string) {
  return /@/.test(message) && (/\btemplate\b/i.test(message) || /\.(?:pptx|xlsx?|csv|docx|pdf|html?|png|jpe?g)\b/i.test(message));
}

export function templateMention(fileName: string) {
  return `@${fileName}`;
}

export function templateOutputFormat(template: ArtifactTemplate): ArtifactFormat {
  if (template.fileType === "pptx") return "pptx";
  if (["xlsx", "xls", "csv"].includes(template.fileType)) return "xlsx";
  if (template.fileType === "docx") return "docx";
  if (template.fileType === "pdf") return "pdf";
  if (["html", "htm"].includes(template.fileType)) return "html";
  return "pptx";
}

export function describeTemplate(template: ArtifactTemplate) {
  return {
    source_id: template.sourceId,
    file_name: template.fileName,
    file_type: template.fileType,
    extraction_status: template.status,
    structural_metadata: template.metadata ?? {},
    extracted_template_content: template.excerpt ?? "No template text was extracted.",
    warnings: template.warnings ?? [],
    semantic_layout_model: template.layoutModel ?? null,
  };
}
