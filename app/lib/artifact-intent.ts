export type ArtifactFormat = "pptx" | "xlsx" | "docx" | "pdf" | "html";

const REQUEST_ACTION = /\b(create|generate|prepare|make|build|produce|turn|draft|develop|put together|export|need|want)\b/i;
const REVISION_ACTION = /\b(change|revise|update|edit|adjust|refine|replace|remove|add|rework|highlight|include|split|make)\b/i;

const FORMAT_PATTERNS: Array<[ArtifactFormat, RegExp]> = [
  ["pptx", /\b(power\s*point|pptx?|presentation|slide\s*deck|deck|slides?)\b/i],
  ["xlsx", /\b(excel|xlsx|spreadsheet|workbook|excel\s+(?:report|dashboard)|tracker\s+in\s+excel)\b/i],
  ["docx", /\b(word|docx|word\s+(?:report|document)|document)\b/i],
  ["pdf", /\b(pdf|pdf\s+report|management\s+pdf)\b/i],
  ["html", /\b(html|html\s+dashboard|web\s+dashboard|standalone\s+dashboard|interactive\s+html\s+report)\b/i],
];

export function detectArtifactRequest(message: string, previousFormat?: ArtifactFormat | null): ArtifactFormat | null {
  for (const [format, pattern] of FORMAT_PATTERNS) {
    if (pattern.test(message) && (REQUEST_ACTION.test(message) || /\b\d+\s+slides?\b/i.test(message))) return format;
  }
  if (previousFormat && REVISION_ACTION.test(message)) return previousFormat;
  return null;
}
