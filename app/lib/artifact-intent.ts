export type ArtifactFormat = "pptx" | "xlsx" | "docx" | "pdf" | "html";

const REQUEST_ACTION = /\b(create|generate|prepare|make|build|produce|turn|draft|develop|put together|export|save|need|want|erstell\w*|generier\w*|bereit\w*|mach\w*|bau\w*|produzier\w*|entwerf\w*|entwickl\w*|exportier\w*|speicher\w*|brauch\w*|möcht\w*|moecht\w*)\b/iu;
const REVISION_ACTION = /\b(change|revise|update|edit|adjust|refine|replace|remove|add|rework|highlight|include|split|make|änder\w*|aender\w*|überarbeit\w*|ueberarbeit\w*|aktualisier\w*|bearbeit\w*|pass\w*|verfeiner\w*|ersetz\w*|entfern\w*|füg\w*|fueg\w*|heb\w*|nutz\w*|verwend\w*|speicher\w*|konvertier\w*)\b/iu;

const FORMAT_PATTERNS: Array<[ArtifactFormat, RegExp]> = [
  ["pptx", /\b(power\s*point|pptx?|presentation|slide\s*deck|deck|slides?|präsentation|praesentation|folien?(?:satz)?)\b/iu],
  ["xlsx", /\b(excel|xlsx|spreadsheet|workbook|excel\s+(?:report|dashboard)|tracker\s+in\s+excel|tabellenkalkulation|arbeitsmappe|excel[- ]?(?:bericht|dashboard))\b/iu],
  ["docx", /\b(word|docx|word\s+(?:report|document)|document|word[- ]?(?:bericht|datei|dokument)|dokument)\b/iu],
  ["pdf", /\b(pdf|pdf\s+report|management\s+pdf|pdf[- ]?bericht)\b/iu],
  ["html", /\b(html|html\s+dashboard|web\s+dashboard|standalone\s+dashboard|interactive\s+html\s+report|web[- ]?dashboard|html[- ]?bericht)\b/iu],
];

export function detectArtifactRequest(message: string, previousFormat?: ArtifactFormat | null): ArtifactFormat | null {
  for (const [format, pattern] of FORMAT_PATTERNS) {
    if (pattern.test(message) && (REQUEST_ACTION.test(message) || /\b\d+\s+(?:slides?|folien?)\b/iu.test(message))) return format;
  }
  if (previousFormat && REVISION_ACTION.test(message)) return previousFormat;
  return null;
}

export function isArtifactRevisionRequest(message: string) {
  return REVISION_ACTION.test(message) && !REQUEST_ACTION.test(message);
}
