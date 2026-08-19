import type { ArtifactFormat } from "./artifact-intent.ts";

export type ExistingContentFormat = Exclude<ArtifactFormat, "xlsx">;
export type ExistingContentBlockKind = "heading" | "paragraph" | "bullet" | "numbered" | "quote" | "code" | "table";

export type ExistingContentRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  code?: boolean;
};

export type ExistingContentTable = {
  headers: string[];
  rows: string[][];
};

export type ExistingContentBlock = {
  id: string;
  kind: ExistingContentBlockKind;
  text: string;
  level?: number;
  locked: boolean;
  runs?: ExistingContentRun[];
  table?: ExistingContentTable;
};

export type ExistingContentMetric = {
  label: string;
  value: string;
  numericValue: number;
  sourceBlockId: string;
};

export type ExistingContentComponent = "title" | "section_heading" | "executive_message" | "body" | "callout" | "native_list" | "native_table" | "kpi_strip" | "two_column";
export type ExistingContentVisualization = "kpi_cards" | "bar_chart" | "timeline";

export type ExistingContentDesignPlan = {
  style: "consulting";
  placements: Array<{
    blockIds: string[];
    component: ExistingContentComponent;
    page: number;
    column: "full" | "left" | "right";
    emphasis: "standard" | "high";
  }>;
  visualizations: Array<{
    type: ExistingContentVisualization;
    sourceBlockIds: string[];
    page: number;
  }>;
};

export type ExistingContentRequest = {
  generationMode: "render_existing_content";
  format: ExistingContentFormat;
  sourceMessage: string;
  blocks: ExistingContentBlock[];
  editableBlockIds: string[];
  editInstruction?: string;
};

const CONTENT_REFERENCE = /\b(?:this text|the text above|text above|above|previous (?:answer|response)|your previous (?:answer|response)|what you just wrote|same content|exactly this|use this content|this content|this)\b/i;
const EDIT_ACTION = /\b(?:shorten|condense|summari[sz]e|rewrite|rephrase|expand|translate|remove|change|edit|revise|adjust)\b/i;

export function referencesExistingContent(message: string) {
  if (/\bthis\s+(?:uploaded\s+)?(?:file|upload|source|data|spreadsheet|workbook|deck)\b/i.test(message) && !/\b(?:this text|this content|previous (?:answer|response)|above)\b/i.test(message)) return false;
  return CONTENT_REFERENCE.test(message);
}

function visibleInlineMarkdown(value: string) {
  return value
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
    .replace(/(\*\*|__)(.*?)\1/g, "$2")
    .replace(/(\*|_)(.*?)\1/g, "$2")
    .replace(/~~(.*?)~~/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\\([\\`*_{}\[\]()#+\-.!>])/g, "$1");
}

export function parseInlineMarkdown(value: string): ExistingContentRun[] {
  const source = value
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1");
  const runs: ExistingContentRun[] = [];
  const pattern = /(\*\*|__)([\s\S]+?)\1|(\*|_)([\s\S]+?)\3|`([^`]+)`/g;
  let cursor = 0;
  for (const match of source.matchAll(pattern)) {
    const index = match.index ?? 0;
    if (index > cursor) runs.push({ text: visibleInlineMarkdown(source.slice(cursor, index)) });
    if (match[2] !== undefined) runs.push({ text: visibleInlineMarkdown(match[2]), bold: true });
    else if (match[4] !== undefined) runs.push({ text: visibleInlineMarkdown(match[4]), italic: true });
    else runs.push({ text: match[5], code: true });
    cursor = index + match[0].length;
  }
  if (cursor < source.length) runs.push({ text: visibleInlineMarkdown(source.slice(cursor)) });
  return runs.filter((run) => run.text.length > 0);
}

function tableCells(line: string) {
  return line.trim().replace(/^\||\|$/g, "").split("|").map((cell) => visibleInlineMarkdown(cell).trim());
}

function isTableSeparator(line: string) {
  const cells = line.trim().replace(/^\||\|$/g, "").split("|");
  return cells.length > 1 && cells.every((cell) => /^\s*:?-{3,}:?\s*$/.test(cell));
}

export function parseExistingContent(markdown: string): ExistingContentBlock[] {
  const blocks: ExistingContentBlock[] = [];
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  let paragraph: string[] = [];
  let inFence = false;
  let code: string[] = [];

  const push = (kind: ExistingContentBlockKind, text: string, level?: number, table?: ExistingContentTable) => {
    const normalized = visibleInlineMarkdown(text).trim();
    if (!normalized) return;
    blocks.push({ id: `block_${blocks.length + 1}`, kind, text: normalized, level, locked: true, runs: kind === "table" ? undefined : parseInlineMarkdown(text), table });
  };
  const flushParagraph = () => {
    if (paragraph.length) push("paragraph", paragraph.join("\n"));
    paragraph = [];
  };
  const flushCode = () => {
    if (code.length) push("code", code.join("\n"));
    code = [];
  };

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    if (/^\s*```/.test(line)) {
      flushParagraph();
      if (inFence) flushCode();
      inFence = !inFence;
      continue;
    }
    if (inFence) {
      code.push(line);
      continue;
    }
    if (line.includes("|") && lineIndex + 1 < lines.length && isTableSeparator(lines[lineIndex + 1])) {
      flushParagraph();
      const headers = tableCells(line);
      const rows: string[][] = [];
      lineIndex += 2;
      while (lineIndex < lines.length && lines[lineIndex].includes("|") && lines[lineIndex].trim()) {
        rows.push(tableCells(lines[lineIndex]));
        lineIndex += 1;
      }
      lineIndex -= 1;
      const table = { headers, rows };
      push("table", [headers, ...rows].flat().join("\n"), undefined, table);
      continue;
    }
    const heading = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/);
    const bullet = line.match(/^\s*[-+*]\s+(.+)$/);
    const numbered = line.match(/^\s*\d+[.)]\s+(.+)$/);
    const quote = line.match(/^\s*>\s?(.*)$/);
    if (heading) {
      flushParagraph();
      push("heading", heading[2], heading[1].length);
    } else if (bullet) {
      flushParagraph();
      push("bullet", bullet[1]);
    } else if (numbered) {
      flushParagraph();
      push("numbered", numbered[1]);
    } else if (quote) {
      flushParagraph();
      push("quote", quote[1]);
    } else if (!line.trim()) {
      flushParagraph();
    } else {
      paragraph.push(line.trim());
    }
  }
  flushParagraph();
  flushCode();
  return blocks;
}

function metricNumber(value: string) {
  const match = value.match(/[+−-]?\s*(?:€|\$|£)?\s*(\d[\d.,]*)/);
  if (!match) return null;
  const normalized = match[1].includes(",") && match[1].includes(".")
    ? match[1].replace(/,/g, "")
    : match[1].replace(",", ".");
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

export function deriveLockedMetrics(blocks: ExistingContentBlock[]): ExistingContentMetric[] {
  const metrics: ExistingContentMetric[] = [];
  const add = (sourceBlockId: string, label: string, value: string) => {
    const numericValue = metricNumber(value);
    if (numericValue === null || metrics.some((metric) => metric.sourceBlockId === sourceBlockId && metric.label === label && metric.value === value)) return;
    metrics.push({ label: label.trim(), value: value.trim(), numericValue, sourceBlockId });
  };
  for (const block of blocks) {
    if (block.table) {
      for (const row of block.table.rows) if (row.length >= 2) add(block.id, row[0], row[1]);
      continue;
    }
    for (const line of block.text.split("\n")) {
      const pair = line.match(/^\s*([^:]{1,80}):\s*([+−-]?(?:(?:€|\$|£)\s*)?\d[\d.,]*(?:\s*%|[kmb])?)\s*$/i);
      if (pair) add(block.id, pair[1], pair[2]);
      const sentencePattern = /\b(synergy target|target|current initiative forecasts total|forecast|gap|shortfall|budget|variance|integration progress)\b[^€$£\d%]{0,28}([+−-]?(?:(?:€|\$|£)\s*)?\d[\d.,]*(?:\s*%|[kmb])?)/gi;
      for (const match of line.matchAll(sentencePattern)) add(block.id, match[1], match[2]);
    }
  }
  return metrics.slice(0, 12);
}

const COMPONENTS: ExistingContentComponent[] = ["title", "section_heading", "executive_message", "body", "callout", "native_list", "native_table", "kpi_strip", "two_column"];
const VISUALIZATIONS: ExistingContentVisualization[] = ["kpi_cards", "bar_chart", "timeline"];

export function defaultExistingContentDesignPlan(blocks: ExistingContentBlock[]): ExistingContentDesignPlan {
  let page = 1;
  let weight = 0;
  const placements: ExistingContentDesignPlan["placements"] = [];
  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index];
    const blockWeight = block.kind === "heading" ? 2 : block.kind === "table" ? 5 : Math.max(1, Math.ceil(block.text.length / 420));
    if (weight > 0 && weight + blockWeight > 9) { page += 1; weight = 0; }
    const previousHeading = [...blocks.slice(0, index)].reverse().find((candidate) => candidate.kind === "heading");
    const executive = previousHeading?.text.match(/executive (?:summary|message)/i);
    const component: ExistingContentComponent = block.kind === "heading"
      ? index === 0 && (block.level ?? 1) === 1 ? "title" : "section_heading"
      : block.kind === "table" ? "native_table"
        : block.kind === "bullet" || block.kind === "numbered" ? "native_list"
          : executive ? "executive_message"
            : block.kind === "quote" ? "callout" : "body";
    placements.push({ blockIds: [block.id], component, page, column: "full", emphasis: component === "executive_message" || component === "callout" ? "high" : "standard" });
    weight += blockWeight;
    if (component === "title" && index < blocks.length - 1) { page += 1; weight = 0; }
  }
  const metrics = deriveLockedMetrics(blocks);
  const metricSources = [...new Set(metrics.map((metric) => metric.sourceBlockId))];
  const visualizations: ExistingContentDesignPlan["visualizations"] = metrics.length
    ? [{ type: metrics.length >= 2 ? "bar_chart" : "kpi_cards", sourceBlockIds: metricSources, page: placements.find((placement) => placement.blockIds.some((id) => metricSources.includes(id)))?.page ?? 1 }]
    : [];
  return { style: "consulting", placements, visualizations };
}

export const existingContentDesignStructuredOutput: { name: string; schema: Record<string, unknown> } = {
  name: "pmi_locked_content_design",
  schema: {
    type: "object",
    additionalProperties: false,
    properties: {
      style: { type: "string", enum: ["consulting"] },
      placements: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            blockIds: { type: "array", items: { type: "string" } },
            component: { type: "string", enum: COMPONENTS },
            page: { type: "integer", minimum: 1, maximum: 50 },
            column: { type: "string", enum: ["full", "left", "right"] },
            emphasis: { type: "string", enum: ["standard", "high"] },
          },
          required: ["blockIds", "component", "page", "column", "emphasis"],
        },
      },
      visualizations: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            type: { type: "string", enum: VISUALIZATIONS },
            sourceBlockIds: { type: "array", items: { type: "string" } },
            page: { type: "integer", minimum: 1, maximum: 50 },
          },
          required: ["type", "sourceBlockIds", "page"],
        },
      },
    },
    required: ["style", "placements", "visualizations"],
  },
};

function jsonObject(raw: string) {
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  return JSON.parse(fenced ?? raw.slice(start, end + 1)) as Record<string, unknown>;
}

function hasOnlyKeys(value: Record<string, unknown>, allowed: string[]) {
  return Object.keys(value).every((key) => allowed.includes(key));
}

export function parseExistingContentDesignPlan(raw: string, blocks: ExistingContentBlock[]): ExistingContentDesignPlan {
  const value = jsonObject(raw);
  if (!hasOnlyKeys(value, ["style", "placements", "visualizations"]) || value.style !== "consulting" || !Array.isArray(value.placements) || !Array.isArray(value.visualizations)) throw new Error("The design plan is incomplete or contains forbidden fields.");
  const knownIds = new Set(blocks.map((block) => block.id));
  const placedIds: string[] = [];
  const placements = value.placements.map((candidate) => {
    if (!candidate || typeof candidate !== "object") throw new Error("The design plan contains an invalid placement.");
    const placement = candidate as Record<string, unknown>;
    if (!hasOnlyKeys(placement, ["blockIds", "component", "page", "column", "emphasis"])) throw new Error("A design placement contains a forbidden text field.");
    const blockIds = Array.isArray(placement.blockIds) ? placement.blockIds.filter((id): id is string => typeof id === "string") : [];
    if (!blockIds.length || blockIds.some((id) => !knownIds.has(id)) || !COMPONENTS.includes(placement.component as ExistingContentComponent)) throw new Error("The design plan references an unknown block or component.");
    if (!Number.isInteger(placement.page) || Number(placement.page) < 1 || !["full", "left", "right"].includes(String(placement.column)) || !["standard", "high"].includes(String(placement.emphasis))) throw new Error("The design plan contains invalid page or placement options.");
    placedIds.push(...blockIds);
    return { blockIds, component: placement.component as ExistingContentComponent, page: Number(placement.page), column: placement.column as "full" | "left" | "right", emphasis: placement.emphasis as "standard" | "high" };
  });
  if (placedIds.length !== knownIds.size || new Set(placedIds).size !== knownIds.size || [...knownIds].some((id) => !placedIds.includes(id))) throw new Error("Every locked block must be placed exactly once.");
  const sourceOrder = blocks.map((block) => block.id);
  if (placedIds.some((id, index) => id !== sourceOrder[index])) throw new Error("Locked blocks must retain their original narrative order.");
  if (placements.some((placement, index) => index > 0 && placement.page < placements[index - 1].page)) throw new Error("Design pages must follow the original narrative order.");
  const metrics = deriveLockedMetrics(blocks);
  const visualizations = value.visualizations.map((candidate) => {
    if (!candidate || typeof candidate !== "object") throw new Error("The design plan contains an invalid visualization.");
    const visualization = candidate as Record<string, unknown>;
    if (!hasOnlyKeys(visualization, ["type", "sourceBlockIds", "page"])) throw new Error("A visualization contains a forbidden authored-data field.");
    const sourceBlockIds = Array.isArray(visualization.sourceBlockIds) ? visualization.sourceBlockIds.filter((id): id is string => typeof id === "string") : [];
    const supportedMetrics = metrics.filter((metric) => sourceBlockIds.includes(metric.sourceBlockId));
    if (!VISUALIZATIONS.includes(visualization.type as ExistingContentVisualization) || !sourceBlockIds.length || sourceBlockIds.some((id) => !knownIds.has(id)) || !supportedMetrics.length) throw new Error("A visualization lacks locked source data.");
    if (visualization.type === "bar_chart" && supportedMetrics.length < 2) throw new Error("A bar chart requires at least two locked metrics.");
    if (!Number.isInteger(visualization.page) || Number(visualization.page) < 1) throw new Error("A visualization has an invalid page.");
    return { type: visualization.type as ExistingContentVisualization, sourceBlockIds, page: Number(visualization.page) };
  });
  if (!visualizations.length && metrics.length) {
    const sourceBlockIds = [...new Set(metrics.map((metric) => metric.sourceBlockId))];
    visualizations.push({
      type: metrics.length >= 2 ? "bar_chart" : "kpi_cards",
      sourceBlockIds,
      page: placements.find((placement) => placement.blockIds.some((id) => sourceBlockIds.includes(id)))?.page ?? 1,
    });
  }
  return { style: "consulting", placements, visualizations };
}

export function buildExistingContentDesignPrompt(input: { blocks: ExistingContentBlock[]; format: ExistingContentFormat; request: string; templateDescription?: unknown }) {
  const blocks = input.blocks.map(({ id, kind, text, level, locked, table }) => ({ id, kind, text, level, locked, table }));
  const metrics = deriveLockedMetrics(input.blocks);
  return `You are a consulting publication designer. Plan the visual representation of approved, immutable content for a ${input.format.toUpperCase()} artifact.

The schema deliberately has no text or replacement_text field. Refer to content only by block ID. Place every block exactly once. You may choose executive-message boxes, native lists/tables, columns, callouts, KPI strips, and source-backed visualizations. Do not invent headings, labels, values, facts, recommendations, or implications. Visualizations may reference only supplied metric source block IDs; the renderer will derive the chart payload deterministically from the locked blocks.

Use professional Deloitte-style hierarchy, page composition, spacing, and management readability. Preserve Markdown semantics as native formatting. A selected template may influence layout, but never wording.

User request: ${input.request}
Format: ${input.format}
Template description: ${JSON.stringify(input.templateDescription ?? null)}
Immutable blocks: ${JSON.stringify(blocks)}
Deterministically supported metric candidates: ${JSON.stringify(metrics)}`;
}

function editTarget(message: string) {
  const match = message.match(/\b(?:the\s+)?([a-z][a-z0-9 /&-]{1,60}?)(?:\s+section)?(?:[.,]|$)/i);
  if (!match) return null;
  const candidate = match[1].replace(/^(?:only\s+)?/, "").trim();
  const known = candidate.match(/(?:executive summary|executive message|summary|conclusion|recommendations?|risks?|actions?|introduction|title|headings?)/i)?.[0];
  return known?.toLowerCase() ?? null;
}

export function editableBlocksForRequest(blocks: ExistingContentBlock[], message: string) {
  if (!EDIT_ACTION.test(message) || /\b(?:do not|don't|without)\s+(?:change|edit|rewrite|shorten|rephrase)/i.test(message)) return [];
  const target = editTarget(message);
  if (!target) return blocks.map((block) => block.id);

  const headingIndex = blocks.findIndex((block) => block.kind === "heading" && block.text.toLowerCase().includes(target));
  if (headingIndex >= 0) {
    const heading = blocks[headingIndex];
    const end = blocks.findIndex((block, index) => index > headingIndex && block.kind === "heading" && (block.level ?? 6) <= (heading.level ?? 6));
    return blocks.slice(headingIndex + 1, end < 0 ? blocks.length : end).map((block) => block.id);
  }

  if (target === "title" || target === "heading" || target === "headings") {
    return blocks.filter((block) => block.kind === "heading").map((block) => block.id);
  }
  if (target.includes("summary") || target === "executive message") {
    const firstNarrative = blocks.find((block) => block.kind === "paragraph" || block.kind === "quote");
    return firstNarrative ? [firstNarrative.id] : [];
  }
  return [];
}

export function resolveExistingContentRequest(input: {
  message: string;
  format: ArtifactFormat;
  history: Array<{ role: "user" | "assistant"; content: string }>;
}): ExistingContentRequest | null {
  if (input.format === "xlsx" || !referencesExistingContent(input.message)) return null;
  const sourceMessage = [...input.history].reverse().find((entry) => entry.role === "assistant" && entry.content.trim())?.content;
  if (!sourceMessage) return null;
  const blocks = parseExistingContent(sourceMessage);
  if (!blocks.length) return null;
  const editableBlockIds = editableBlocksForRequest(blocks, input.message);
  const editable = new Set(editableBlockIds);
  return {
    generationMode: "render_existing_content",
    format: input.format,
    sourceMessage,
    blocks: blocks.map((block) => ({ ...block, locked: !editable.has(block.id) })),
    editableBlockIds,
    editInstruction: editableBlockIds.length ? input.message : undefined,
  };
}

export function applyBlockEdits(blocks: ExistingContentBlock[], edits: Array<{ id: string; text: string }>) {
  const editableIds = new Set(blocks.filter((block) => !block.locked).map((block) => block.id));
  if (edits.length !== editableIds.size) throw new Error("The edit response did not return every editable text block exactly once.");
  const replacements = new Map<string, string>();
  for (const edit of edits) {
    if (!editableIds.has(edit.id) || replacements.has(edit.id) || typeof edit.text !== "string" || !edit.text.trim()) {
      throw new Error("The edit response attempted to change a locked block or returned an invalid replacement.");
    }
    replacements.set(edit.id, edit.text.trim());
  }
  return blocks.map((block) => {
    const replacement = replacements.get(block.id);
    if (replacement === undefined) return { ...block };
    const parsed = parseExistingContent(replacement);
    const semantic = parsed.length === 1 ? parsed[0] : null;
    return {
      ...block,
      kind: semantic?.kind ?? block.kind,
      text: semantic?.text ?? visibleInlineMarkdown(replacement).trim(),
      runs: semantic?.runs ?? parseInlineMarkdown(replacement),
      table: semantic?.table,
      locked: false,
    };
  });
}

export function normalizedText(value: string) {
  return visibleInlineMarkdown(value)
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\s*\n\s*/g, " ")
    .trim();
}

export function assertLockedBlocksUnchanged(original: ExistingContentBlock[], final: ExistingContentBlock[]) {
  const byId = new Map(final.map((block) => [block.id, block]));
  for (const block of original.filter((candidate) => candidate.locked)) {
    const rendered = byId.get(block.id);
    if (!rendered || normalizedText(rendered.text) !== normalizedText(block.text)) {
      throw new Error(`Locked text integrity check failed for ${block.id}.`);
    }
  }
}

export function assertRenderedTextIntegrity(expected: ExistingContentBlock[], renderedTextBlocks: string[]) {
  const expectedText = expected.map((block) => normalizedText(block.text));
  const actualText = renderedTextBlocks.map(normalizedText);
  if (expectedText.length !== actualText.length || expectedText.some((text, index) => text !== actualText[index])) {
    throw new Error("Rendered text integrity check failed: the file content differs from the locked textual blocks.");
  }
}

export function assertVisualizationDataIntegrity(blocks: ExistingContentBlock[], plan: ExistingContentDesignPlan) {
  const byId = new Map(blocks.map((block) => [block.id, normalizedText(block.text)]));
  const metrics = deriveLockedMetrics(blocks);
  for (const visualization of plan.visualizations) {
    const selected = metrics.filter((metric) => visualization.sourceBlockIds.includes(metric.sourceBlockId));
    if (!selected.length) throw new Error("Visualization integrity check failed: no locked metric supports the visual.");
    for (const metric of selected) {
      const source = byId.get(metric.sourceBlockId) ?? "";
      if (!source.includes(normalizedText(metric.label)) || !source.includes(normalizedText(metric.value))) {
        throw new Error("Visualization integrity check failed: a label or value is not present in its locked source block.");
      }
    }
  }
}

export const blockEditStructuredOutput: { name: string; description: string; schema: Record<string, unknown> } = {
  name: "pmi_locked_text_edits",
  description: "Return replacements only for the explicitly editable text blocks.",
  schema: {
    type: "object",
    additionalProperties: false,
    properties: {
      blocks: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: { id: { type: "string" }, text: { type: "string" } },
          required: ["id", "text"],
        },
      },
    },
    required: ["blocks"],
  },
};

export function parseBlockEditResponse(raw: string) {
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  const value = JSON.parse(fenced ?? raw.slice(start, end + 1)) as { blocks?: unknown };
  if (!Array.isArray(value.blocks)) throw new Error("The edit response did not contain a blocks array.");
  return value.blocks.map((entry) => {
    if (!entry || typeof entry !== "object") throw new Error("The edit response contained an invalid block.");
    const block = entry as Record<string, unknown>;
    if (typeof block.id !== "string" || typeof block.text !== "string") throw new Error("The edit response contained an invalid block ID or text.");
    return { id: block.id, text: block.text };
  });
}

export function buildScopedEditPrompt(blocks: ExistingContentBlock[], instruction: string) {
  const editable = blocks.filter((block) => !block.locked).map(({ id, text }) => ({ id, text }));
  return `Apply only the user's explicit edit to the supplied editable text blocks. Do not add analysis, recommendations, headings, or source material. Return exactly one replacement for every supplied ID and no other IDs.\n\nInstruction: ${instruction}\nEditable blocks: ${JSON.stringify(editable)}`;
}
