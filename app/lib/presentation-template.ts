import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { PresentationItem, PresentationModel, PresentationSlide } from "./presentation";

type Bounds = { x: number; y: number; width: number; height: number };
type RegionRole = "title" | "subtitle" | "card" | "panel" | "table" | "takeaway" | "body" | "source" | "footer";

export type PresentationTemplateRegion = {
  id: string;
  role: RegionRole;
  kind: "text" | "card" | "panel" | "table" | "shape-table";
  shapeIds: string[];
  bounds: Bounds;
  fontSize: number;
  maxChars: number;
  maxItems: number;
  textHint: string;
  rows?: number;
  columns?: number;
  headers?: string[];
  statusShapeIds?: string[];
};

export type PresentationTemplateSlide = {
  slideNumber: number;
  tags: string[];
  regions: PresentationTemplateRegion[];
  maxItems: number;
};

export type PresentationTemplateLayout = {
  kind: "pptx";
  slideWidth: number;
  slideHeight: number;
  slides: PresentationTemplateSlide[];
};

type ShapeInfo = {
  id: string;
  name: string;
  xml: string;
  text: string;
  bounds: Bounds;
  fontSize: number;
  placeholder: string;
  maxChars: number;
};

type TableInfo = {
  id: string;
  xml: string;
  bounds: Bounds;
  rows: string[][];
  columnWidths: number[];
};

type SlideAnalysis = {
  model: PresentationTemplateSlide;
  shapes: ShapeInfo[];
  tables: TableInfo[];
};

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const NOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const EMU_PER_INCH = 914400;

function xml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function plain(value: string) {
  return value.replace(/<[^>]+>/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&").replace(/\s+/g, " ").trim();
}

function textNodes(value: string) {
  return [...value.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((match) => plain(match[1])).filter(Boolean);
}

function shapeText(shape: string) {
  return textNodes(shape).join(" ");
}

function shapeId(shape: string, fallback: string) {
  return shape.match(/<p:cNvPr\b[^>]*\bid="([^"]+)"/)?.[1] ?? fallback;
}

function shapeName(shape: string) {
  return plain(shape.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1] ?? "");
}

function placeholderType(shape: string) {
  return shape.match(/<p:ph\b[^>]*\btype="([^"]+)"/)?.[1] ?? "";
}

function boundsOf(value: string): Bounds {
  const transform = value.match(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/)?.[0] ?? value;
  const x = Number(transform.match(/<a:off\b[^>]*\bx="(\d+)"/)?.[1] ?? 0);
  const y = Number(transform.match(/<a:off\b[^>]*\by="(\d+)"/)?.[1] ?? 0);
  const width = Number(transform.match(/<a:ext\b[^>]*\bcx="(\d+)"/)?.[1] ?? 0);
  const height = Number(transform.match(/<a:ext\b[^>]*\bcy="(\d+)"/)?.[1] ?? 0);
  return { x, y, width, height };
}

function fontSizeOf(value: string) {
  const sizes = [...value.matchAll(/\bsz="(\d+)"/g)].map((match) => Number(match[1]) / 100).filter((size) => size >= 5 && size <= 200);
  return sizes.length ? Math.max(...sizes) : 12;
}

function textCapacity(bounds: Bounds, fontSize: number) {
  if (!bounds.width || !bounds.height) return 100;
  const widthPoints = bounds.width / EMU_PER_INCH * 72;
  const heightPoints = bounds.height / EMU_PER_INCH * 72;
  const charsPerLine = Math.max(5, Math.floor(widthPoints / Math.max(4, fontSize * 0.53)));
  const lines = Math.max(1, Math.floor(heightPoints / Math.max(8, fontSize * 1.2)));
  return Math.max(12, Math.floor(charsPerLine * lines * 0.86));
}

function contains(outer: Bounds, inner: Bounds) {
  if (!outer.width || !outer.height || !inner.width || !inner.height) return false;
  const centerX = inner.x + inner.width / 2;
  const centerY = inner.y + inner.height / 2;
  return centerX >= outer.x && centerX <= outer.x + outer.width && centerY >= outer.y && centerY <= outer.y + outer.height;
}

function area(bounds: Bounds) {
  return bounds.width * bounds.height;
}

function unionBounds(values: Bounds[]): Bounds {
  if (!values.length) return { x: 0, y: 0, width: 0, height: 0 };
  const x = Math.min(...values.map((value) => value.x));
  const y = Math.min(...values.map((value) => value.y));
  const right = Math.max(...values.map((value) => value.x + value.width));
  const bottom = Math.max(...values.map((value) => value.y + value.height));
  return { x, y, width: right - x, height: bottom - y };
}

function detectShapeTable(shapes: ShapeInfo[], slideWidth: number, slideHeight: number) {
  const candidates = shapes.filter((shape) => shape.bounds.y > slideHeight * 0.24 && shape.bounds.y < slideHeight * 0.78 && shape.bounds.x < slideWidth * 0.72);
  const tolerance = slideHeight * 0.028;
  const rows: ShapeInfo[][] = [];
  for (const shape of [...candidates].sort((a, b) => a.bounds.y - b.bounds.y || a.bounds.x - b.bounds.x)) {
    const center = shape.bounds.y + shape.bounds.height / 2;
    const row = rows.find((entry) => Math.abs((entry[0].bounds.y + entry[0].bounds.height / 2) - center) <= tolerance);
    if (row) row.push(shape); else rows.push([shape]);
  }
  const eligible = rows.map((row) => row.sort((a, b) => a.bounds.x - b.bounds.x)).filter((row) => row.length >= 3);
  let best: ShapeInfo[][] = [];
  for (let start = 0; start < eligible.length; start++) {
    const sequence = [eligible[start]];
    for (let index = start + 1; index < eligible.length; index++) {
      const previousY = sequence.at(-1)?.[0].bounds.y ?? 0;
      if (eligible[index][0].bounds.y - previousY > slideHeight * 0.12) break;
      const columns = Math.min(sequence[0].length, eligible[index].length);
      const aligned = Array.from({ length: columns }, (_, column) => Math.abs(sequence[0][column].bounds.x - eligible[index][column].bounds.x) < slideWidth * 0.04).filter(Boolean).length;
      if (aligned < 3) break;
      sequence.push(eligible[index]);
    }
    if (sequence.length > best.length) best = sequence;
  }
  if (best.length < 4) return null;
  const columnCount = Math.min(...best.map((row) => row.length));
  const normalizedRows = best.map((row) => row.slice(0, columnCount));
  return {
    rows: normalizedRows,
    headers: normalizedRows[0].map((shape) => shape.text),
    bounds: unionBounds(normalizedRows.flat().map((shape) => shape.bounds)),
  };
}

function extractTables(slideXml: string) {
  return [...slideXml.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)]
    .map((match, index): TableInfo | null => {
      const frame = match[0];
      if (!/<a:tbl>/.test(frame)) return null;
      const rows = [...frame.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((row) => [...row[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((cell) => textNodes(cell[0]).join(" ")));
      return {
        id: shapeId(frame, `table-${index + 1}`),
        xml: frame,
        bounds: boundsOf(frame),
        rows,
        columnWidths: [...frame.matchAll(/<a:gridCol\b[^>]*\bw="(\d+)"/g)].map((column) => Number(column[1])),
      };
    })
    .filter((table): table is TableInfo => Boolean(table));
}

function inferTags(text: string) {
  const tags: string[] = [];
  const rules: Array<[string, RegExp]> = [
    ["decisions", /\b(decision|approve|approval|steerco)\b/i],
    ["risks", /\b(risk|issue|mitigation|red|amber)\b/i],
    ["synergies", /\b(synerg|value capture|benefit)\b/i],
    ["timeline", /\b(milestone|timeline|roadmap|date)\b/i],
    ["summary", /\b(summary|overview|executive)\b/i],
  ];
  for (const [tag, pattern] of rules) if (pattern.test(text)) tags.push(tag);
  return tags;
}

function isGenericChrome(value: string) {
  return /^(?:confidential|strictly confidential|private and confidential|sources?|source notes?|appendix|page \d+)$/i.test(value)
    || /(?:copyright|all rights reserved|©)/i.test(value);
}

function analyzeSlide(slideXml: string, slideNumber: number, slideWidth: number, slideHeight: number): SlideAnalysis {
  const shapeMatches = [...slideXml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)];
  const shapes = shapeMatches.map((match, index): ShapeInfo => {
    const shape = match[0];
    const bounds = boundsOf(shape);
    const fontSize = fontSizeOf(shape);
    return {
      id: shapeId(shape, `shape-${index + 1}`),
      name: shapeName(shape),
      xml: shape,
      text: shapeText(shape),
      bounds,
      fontSize,
      placeholder: placeholderType(shape),
      maxChars: textCapacity(bounds, fontSize),
    };
  });
  const tables = extractTables(slideXml);
  const textShapes = shapes.filter((shape) => shape.text && !["sldNum", "dt", "ftr"].includes(shape.placeholder));
  const topCandidates = textShapes.filter((shape) => !shape.bounds.y || shape.bounds.y < slideHeight * 0.35);
  const formalTitle = textShapes.find((shape) => shape.placeholder === "title" || shape.placeholder === "ctrTitle");
  const title = formalTitle ?? [...topCandidates].sort((a, b) => b.fontSize - a.fontSize || a.bounds.y - b.bounds.y)[0];
  const subtitle = [...topCandidates]
    .filter((shape) => shape.id !== title?.id && shape.bounds.y >= (title?.bounds.y ?? 0))
    .sort((a, b) => a.bounds.y - b.bounds.y || b.fontSize - a.fontSize)[0];

  const containerShapes = shapes.filter((shape) => !shape.text && shape.bounds.width && shape.bounds.height && area(shape.bounds) < slideWidth * slideHeight * 0.65);
  const claimed = new Set<string>([title?.id, subtitle?.id].filter((value): value is string => Boolean(value)));
  const complexContainerChildren = new Set(containerShapes.flatMap((container) => {
    const children = textShapes.filter((shape) => contains(container.bounds, shape.bounds));
    return children.length >= 4 ? children.map((shape) => shape.id) : [];
  }));
  const shapeTable = detectShapeTable(textShapes.filter((shape) => !complexContainerChildren.has(shape.id)), slideWidth, slideHeight);
  if (shapeTable) shapeTable.rows.flat().forEach((shape) => claimed.add(shape.id));
  const cardGroups: Array<{ container?: ShapeInfo; children: ShapeInfo[] }> = [];
  for (const container of [...containerShapes].sort((a, b) => area(a.bounds) - area(b.bounds))) {
    const children = textShapes.filter((shape) => !claimed.has(shape.id) && contains(container.bounds, shape.bounds));
    if (children.length < 2) continue;
    children.forEach((shape) => claimed.add(shape.id));
    cardGroups.push({ container, children });
  }
  for (const shape of textShapes) {
    if (claimed.has(shape.id) || isGenericChrome(shape.text)) continue;
    const selfContained = /<a:solidFill>|<a:gradFill>/.test(shape.xml) && shape.bounds.width < slideWidth * 0.55 && shape.bounds.height > slideHeight * 0.08;
    if (selfContained) {
      claimed.add(shape.id);
      cardGroups.push({ children: [shape] });
    }
  }

  const regions: PresentationTemplateRegion[] = [];
  if (title) regions.push({ id: `title-${title.id}`, role: "title", kind: "text", shapeIds: [title.id], bounds: title.bounds, fontSize: title.fontSize, maxChars: title.maxChars, maxItems: 0, textHint: title.text.slice(0, 120) });
  if (subtitle) regions.push({ id: `subtitle-${subtitle.id}`, role: "subtitle", kind: "text", shapeIds: [subtitle.id], bounds: subtitle.bounds, fontSize: subtitle.fontSize, maxChars: subtitle.maxChars, maxItems: 0, textHint: subtitle.text.slice(0, 120) });

  cardGroups.forEach((group, index) => {
    const ordered = [...group.children].sort((a, b) => a.bounds.y - b.bounds.y || b.fontSize - a.fontSize);
    const bounds = group.container?.bounds ?? unionBounds(ordered.map((shape) => shape.bounds));
    const role: RegionRole = bounds.y > slideHeight * 0.72 && bounds.width > slideWidth * 0.45 ? "takeaway" : ordered.length >= 5 ? "panel" : "card";
    regions.push({
      id: `${role}-${index + 1}`,
      role,
      kind: role === "panel" ? "panel" : role === "takeaway" ? "text" : "card",
      shapeIds: ordered.map((shape) => shape.id),
      bounds,
      fontSize: Math.max(...ordered.map((shape) => shape.fontSize)),
      maxChars: ordered.reduce((sum, shape) => sum + shape.maxChars, 0),
      maxItems: role === "panel" ? Math.floor((ordered.length - 1) / 2) : role === "card" ? 1 : 0,
      textHint: ordered.map((shape) => shape.text).join(" | ").slice(0, 180),
    });
  });

  if (shapeTable) {
    const statusColumn = shapeTable.headers.findIndex((header) => /status|rag|traffic/i.test(header));
    const statusShapeIds = statusColumn < 0 ? [] : shapeTable.rows.slice(1).map((row) => {
      const cell = row[statusColumn];
      const centerY = cell.bounds.y + cell.bounds.height / 2;
      return shapes.filter((shape) => !shape.text && (/oval|ellipse/i.test(shape.name) || (
        shape.bounds.width < slideWidth * 0.03
        && shape.bounds.height < slideHeight * 0.05
        && shape.bounds.width / Math.max(1, shape.bounds.height) > 0.55
        && shape.bounds.width / Math.max(1, shape.bounds.height) < 1.8
      )))
        .sort((a, b) => {
          const distanceA = Math.abs((a.bounds.y + a.bounds.height / 2) - centerY) + Math.abs((a.bounds.x + a.bounds.width / 2) - cell.bounds.x);
          const distanceB = Math.abs((b.bounds.y + b.bounds.height / 2) - centerY) + Math.abs((b.bounds.x + b.bounds.width / 2) - cell.bounds.x);
          return distanceA - distanceB;
        })[0]?.id ?? "";
    });
    regions.push({
      id: "shape-table-1",
      role: "table",
      kind: "shape-table",
      shapeIds: shapeTable.rows.flat().map((shape) => shape.id),
      bounds: shapeTable.bounds,
      fontSize: Math.max(...shapeTable.rows.flat().map((shape) => shape.fontSize)),
      maxChars: 0,
      maxItems: shapeTable.rows.length - 1,
      textHint: shapeTable.headers.join(" | ").slice(0, 180),
      rows: shapeTable.rows.length,
      columns: shapeTable.headers.length,
      headers: shapeTable.headers,
      statusShapeIds,
    });
  }

  tables.forEach((table, index) => {
    const headers = table.rows[0] ?? [];
    regions.push({
      id: `table-${table.id || index + 1}`,
      role: "table",
      kind: "table",
      shapeIds: [table.id],
      bounds: table.bounds,
      fontSize: fontSizeOf(table.xml),
      maxChars: 0,
      maxItems: Math.max(0, table.rows.length - 1),
      textHint: headers.join(" | ").slice(0, 180),
      rows: table.rows.length,
      columns: Math.max(0, ...table.rows.map((row) => row.length)),
      headers,
    });
  });

  const unclaimed = textShapes.filter((shape) => !claimed.has(shape.id) && shape.id !== title?.id && shape.id !== subtitle?.id);
  for (const shape of unclaimed) {
    const bottom = shape.bounds.y + shape.bounds.height;
    const role: RegionRole = /source|caveat|reference/i.test(shape.text) ? "source"
      : shape.placeholder === "ftr" || bottom > slideHeight * 0.94 ? "footer"
        : shape.bounds.y > slideHeight * 0.7 && shape.bounds.width > slideWidth * 0.45 ? "takeaway"
          : "body";
    regions.push({ id: `${role}-${shape.id}`, role, kind: "text", shapeIds: [shape.id], bounds: shape.bounds, fontSize: shape.fontSize, maxChars: shape.maxChars, maxItems: role === "body" ? 1 : 0, textHint: shape.text.slice(0, 120) });
  }

  const maxItems = Math.max(1, regions.filter((region) => region.role === "card" || region.role === "panel" || region.role === "body" || region.role === "table").reduce((sum, region) => sum + region.maxItems, 0));
  const allText = [shapeText(slideXml), ...tables.flatMap((table) => table.rows.flat())].join(" ");
  return { model: { slideNumber, tags: inferTags(allText), regions, maxItems }, shapes, tables };
}

function slideDimensions(files: Record<string, Uint8Array>) {
  const presentation = strFromU8(files["ppt/presentation.xml"]);
  const match = presentation.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  return { width: Number(match?.[1] ?? 12192000), height: Number(match?.[2] ?? 6858000) };
}

function sortedSlidePaths(files: Record<string, Uint8Array>) {
  return Object.keys(files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
}

export function analyzePresentationTemplate(templateBytes: Uint8Array): PresentationTemplateLayout {
  const files = unzipSync(templateBytes);
  if (!files["[Content_Types].xml"] || !files["ppt/presentation.xml"]) throw new Error("The selected PPTX template is not a valid PowerPoint package.");
  const paths = sortedSlidePaths(files);
  if (!paths.length) throw new Error("The selected PPTX template has no slides.");
  const dimensions = slideDimensions(files);
  return {
    kind: "pptx",
    slideWidth: dimensions.width,
    slideHeight: dimensions.height,
    slides: paths.map((path, index) => analyzeSlide(strFromU8(files[path]), index + 1, dimensions.width, dimensions.height).model),
  };
}

function truncateToCapacity(value: string, maxChars: number) {
  const normalized = value.replace(/\s+/g, " ").trim();
  if (normalized.length <= maxChars) return normalized;
  const candidate = normalized.slice(0, Math.max(1, maxChars - 1));
  const wordBoundary = candidate.lastIndexOf(" ");
  return `${candidate.slice(0, wordBoundary > maxChars * 0.55 ? wordBoundary : candidate.length).trim()}…`;
}

function replaceText(shape: string, value: string) {
  let first = true;
  return shape.replace(/<a:t>([\s\S]*?)<\/a:t>/g, () => {
    if (!first) return "<a:t></a:t>";
    first = false;
    return `<a:t>${xml(value)}</a:t>`;
  });
}

function replaceFontSize(shape: string, fontSize: number) {
  const size = Math.round(fontSize * 100);
  return shape.replace(/\bsz="\d+"/g, `sz="${size}"`);
}

function fitSingleLine(value: string, shape: ShapeInfo) {
  const normalized = value.replace(/\s+/g, " ").trim();
  if (normalized.length <= shape.maxChars) return { text: normalized };
  const fontSize = Math.max(8, Math.floor(shape.fontSize * shape.maxChars / normalized.length * 2) / 2);
  const adjustedCapacity = Math.floor(shape.maxChars * shape.fontSize / fontSize);
  return { text: truncateToCapacity(normalized, adjustedCapacity), fontSize };
}

function compactItem(item: PresentationItem) {
  return [item.value, item.label, item.detail, item.implication ? `Implication: ${item.implication}` : undefined].filter(Boolean).join(" — ");
}

function roleValue(role: string, item: PresentationItem) {
  if (/status/i.test(role)) return (item.status ?? "neutral").toUpperCase();
  if (/owner/i.test(role)) return item.owner ?? "TBC";
  if (/\bdate\b|deadline|when|due/i.test(role)) return item.deadline ?? "TBC";
  if (/source|caveat|reference/i.test(role)) return (item.sourceRefs ?? []).join(", ") || "Not evidenced";
  if (/implication|why|impact/i.test(role)) return item.implication ?? item.detail ?? "Not evidenced";
  if (/progress|value|metric/i.test(role)) return item.value ?? (item.status ?? "neutral").toUpperCase();
  if (/update|mitigation|detail|rationale|comment|evidence|description/i.test(role)) return [item.detail, item.implication ? `Implication: ${item.implication}` : undefined].filter(Boolean).join(" ") || "Not evidenced";
  return item.label;
}

function tableCellCapacity(table: TableInfo, column: number, rowCount: number, fontSize: number) {
  const width = table.columnWidths[column] ?? table.bounds.width / Math.max(1, table.rows[0]?.length ?? 1);
  const height = table.bounds.height / Math.max(1, rowCount);
  return textCapacity({ x: 0, y: 0, width, height }, fontSize);
}

function fillTable(frame: string, table: TableInfo, items: PresentationItem[]) {
  const rows = [...frame.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((match) => match[0]);
  const headers = table.rows[0] ?? [];
  let dataIndex = 0;
  let rowNumber = 0;
  return frame.replace(/<a:tr\b[\s\S]*?<\/a:tr>/g, (row) => {
    if (rowNumber++ === 0) return row;
    const item = items[dataIndex++];
    let column = 0;
    return row.replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (cell) => {
      const header = headers[column] ?? `column ${column + 1}`;
      const value = item ? roleValue(header, item) : "";
      const capacity = tableCellCapacity(table, column, rows.length, fontSizeOf(cell));
      column += 1;
      return replaceText(cell, truncateToCapacity(value, capacity));
    });
  });
}

function replaceShapeFill(shape: string, color: string) {
  return shape.replace(/(<p:spPr\b[\s\S]*?<a:solidFill>)[\s\S]*?(<\/a:solidFill>)/, `$1<a:srgbClr val="${color}"/>$2`);
}

function statusColor(status?: PresentationItem["status"]) {
  if (status === "red") return "D92D20";
  if (status === "amber") return "F59E0B";
  if (status === "green") return "2E9B57";
  return "8A8F8C";
}

function chooseSourceSlide(slide: PresentationSlide, analyses: SlideAnalysis[]) {
  if (slide.templateSlide && analyses[slide.templateSlide - 1]) return slide.templateSlide - 1;
  const desired = slide.layout ?? "cards";
  const scores = analyses.map((analysis, index) => {
    let score = analysis.model.tags.includes(desired) ? 10 : 0;
    if (desired === "cover" && index === 0) score += 8;
    if (analysis.model.maxItems >= slide.items.length) score += 4;
    score -= Math.abs(analysis.model.maxItems - slide.items.length) * 0.1;
    return score;
  });
  return scores.indexOf(Math.max(...scores));
}

function expandSlides(model: PresentationModel, analyses: SlideAnalysis[]) {
  return model.slides.flatMap((slide) => {
    const sourceIndex = chooseSourceSlide(slide, analyses);
    const capacity = Math.max(1, analyses[sourceIndex].model.maxItems);
    if (slide.items.length <= capacity) return [{ slide, sourceIndex }];
    const pageCount = Math.ceil(slide.items.length / capacity);
    const itemsPerPage = Math.ceil(slide.items.length / pageCount);
    return Array.from({ length: pageCount }, (_, pageIndex) => ({
      sourceIndex,
      slide: {
        ...slide,
        title: `${slide.title} — ${pageIndex + 1}/${pageCount}`,
        items: slide.items.slice(pageIndex * itemsPerPage, (pageIndex + 1) * itemsPerPage),
      },
    }));
  });
}

function populateSlide(sourceXml: string, analysis: SlideAnalysis, slide: PresentationSlide, model: PresentationModel) {
  const replacements = new Map<string, string>();
  const fillReplacements = new Map<string, string>();
  const fontReplacements = new Map<string, number>();
  const shapeById = new Map(analysis.shapes.map((shape) => [shape.id, shape]));
  const title = analysis.model.regions.find((region) => region.role === "title");
  const subtitle = analysis.model.regions.find((region) => region.role === "subtitle");
  if (title) {
    const shape = shapeById.get(title.shapeIds[0]);
    const fitted = shape ? fitSingleLine(slide.title || model.title, shape) : { text: truncateToCapacity(slide.title || model.title, title.maxChars) };
    replacements.set(title.shapeIds[0], fitted.text);
    if (fitted.fontSize) fontReplacements.set(title.shapeIds[0], fitted.fontSize);
  }
  if (subtitle) replacements.set(subtitle.shapeIds[0], truncateToCapacity(slide.keyMessage ?? slide.kicker ?? model.executiveSummary, subtitle.maxChars));

  let itemIndex = 0;
  const cards = analysis.model.regions.filter((region) => region.role === "card");
  for (const card of cards) {
    const item = slide.items[itemIndex++];
    const editableShapeIds = card.shapeIds.filter((id) => !isGenericChrome(shapeById.get(id)?.text ?? "") && !/^(?:decision required|action required|risk|status|implication|recommendation|management takeaway)$/i.test(shapeById.get(id)?.text ?? ""));
    if (!item) {
      editableShapeIds.forEach((id) => replacements.set(id, ""));
      continue;
    }
    if (editableShapeIds.length <= 1) {
      const id = editableShapeIds[0] ?? card.shapeIds[0];
      replacements.set(id, truncateToCapacity(compactItem(item), shapeById.get(id)?.maxChars ?? card.maxChars));
      continue;
    }
    const ordered = editableShapeIds.map((id) => shapeById.get(id)).filter((shape): shape is ShapeInfo => Boolean(shape)).sort((a, b) => b.fontSize - a.fontSize || a.bounds.y - b.bounds.y);
    const values = slide.layout === "decisions"
      ? [item.label, "DECISION REQUIRED", ""]
      : slide.layout === "risks"
        ? [item.label, `${(item.status ?? "neutral").toUpperCase()} RISK`, ""]
        : [item.value ?? item.label, item.value ? item.label : item.detail ?? "Not evidenced", item.implication ? `Implication: ${item.implication}` : ""];
    ordered.forEach((shape, index) => replacements.set(shape.id, truncateToCapacity(values[index] ?? "", shape.maxChars)));
  }

  const tableRegions = analysis.model.regions.filter((region) => region.role === "table");
  const tableItems = slide.items.slice(itemIndex);
  const tableById = new Map(analysis.tables.map((table) => [table.id, table]));
  const tableReplacements = new Map<string, PresentationItem[]>();
  for (const region of tableRegions) {
    const items = tableItems.splice(0, region.maxItems);
    if (region.kind === "shape-table") {
      const columns = region.columns ?? region.headers?.length ?? 1;
      const headers = region.headers ?? [];
      const dataShapeIds = region.shapeIds.slice(columns);
      for (let row = 0; row < region.maxItems; row++) {
        const item = items[row];
        const statusShapeId = region.statusShapeIds?.[row];
        if (item && statusShapeId) fillReplacements.set(statusShapeId, statusColor(item.status));
        for (let column = 0; column < columns; column++) {
          const id = dataShapeIds[row * columns + column];
          const shape = shapeById.get(id);
          if (!id || !shape) continue;
          const value = item ? roleValue(headers[column] ?? `column ${column + 1}`, item) : "";
          replacements.set(id, truncateToCapacity(value, shape.maxChars));
        }
      }
    } else {
      tableReplacements.set(region.shapeIds[0], items);
    }
  }

  const panels = analysis.model.regions.filter((region) => region.role === "panel");
  for (const panel of panels) {
    const [headingId, ...slotIds] = panel.shapeIds;
    if (/example|nike|adidas/i.test(shapeById.get(headingId)?.text ?? "")) replacements.set(headingId, "Key management insights");
    const panelItems = tableItems.splice(0, panel.maxItems);
    for (let index = 0; index < panel.maxItems; index++) {
      const item = panelItems[index];
      const labelId = slotIds[index * 2];
      const detailId = slotIds[index * 2 + 1];
      const labelShape = shapeById.get(labelId);
      const detailShape = shapeById.get(detailId);
      if (labelId && labelShape) replacements.set(labelId, item ? truncateToCapacity(item.label, labelShape.maxChars) : "");
      if (detailId && detailShape) replacements.set(detailId, item ? truncateToCapacity([item.detail, item.implication ? `Implication: ${item.implication}` : undefined].filter(Boolean).join(" ") || "Not evidenced", detailShape.maxChars) : "");
    }
  }

  const bodyRegions = analysis.model.regions.filter((region) => region.role === "body");
  for (const region of bodyRegions) {
    const item = tableItems.shift();
    replacements.set(region.shapeIds[0], item ? truncateToCapacity(compactItem(item), region.maxChars) : "");
  }
  const takeaway = analysis.model.regions.find((region) => region.role === "takeaway");
  if (takeaway) replacements.set(takeaway.shapeIds[0], truncateToCapacity(slide.keyMessage ?? model.executiveSummary, takeaway.maxChars));
  const sourceRegion = analysis.model.regions.find((region) => region.role === "source");
  if (sourceRegion) replacements.set(sourceRegion.shapeIds[0], truncateToCapacity((slide.sourceNotes ?? []).join(" · "), sourceRegion.maxChars));
  for (const footer of analysis.model.regions.filter((region) => region.role === "footer")) {
    const shape = shapeById.get(footer.shapeIds[0]);
    if (shape && /illustrative|fictional|example/i.test(shape.text)) replacements.set(shape.id, truncateToCapacity("Source-backed PMI summary", shape.maxChars));
  }

  let populated = sourceXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape, index) => {
    const id = shapeId(shape, `shape-${index + 1}`);
    const withText = replacements.has(id) ? replaceText(shape, replacements.get(id) ?? "") : shape;
    const withFont = fontReplacements.has(id) ? replaceFontSize(withText, fontReplacements.get(id) ?? 8) : withText;
    return fillReplacements.has(id) ? replaceShapeFill(withFont, fillReplacements.get(id) ?? "8A8F8C") : withFont;
  });
  populated = populated.replace(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g, (frame, index) => {
    const id = shapeId(frame, `table-${index + 1}`);
    const items = tableReplacements.get(id);
    const table = tableById.get(id);
    return items && table ? fillTable(frame, table, items) : frame;
  });
  return populated;
}

function editNotes(sourceXml: string, slide: PresentationSlide) {
  const sources = [...new Set(slide.items.flatMap((item) => item.sourceRefs ?? []))];
  const note = `[Sources] ${[...sources.map((id) => `Source ID: ${id}`), ...(slide.sourceNotes ?? [])].join("; ") || "No source reference supplied; validate claims before circulation."}`;
  let written = false;
  return sourceXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    if (written || !/<p:ph\b[^>]*\btype="body"/.test(shape) || !/<a:t>/.test(shape)) return shape;
    written = true;
    return replaceText(shape, note);
  });
}

function insertBeforeClosing(value: string, closingTag: string, addition: string) {
  const index = value.lastIndexOf(closingTag);
  if (index < 0) throw new Error(`The selected PPTX template is missing ${closingTag}.`);
  return `${value.slice(0, index)}${addition}${value.slice(index)}`;
}

function addContentType(value: string, partName: string, contentType: string) {
  if (value.includes(`PartName="${partName}"`)) return value;
  return insertBeforeClosing(value, "</Types>", `<Override PartName="${partName}" ContentType="${contentType}"/>`);
}

function notesTarget(rels: string) {
  for (const match of rels.matchAll(/<Relationship\b[^>]*\/?>(?:<\/Relationship>)?/g)) {
    if (!/\bType="[^"]+\/notesSlide"/.test(match[0])) continue;
    return match[0].match(/\bTarget="\.\.\/notesSlides\/notesSlide(\d+)\.xml"/)?.[1];
  }
  return undefined;
}

function validatePopulatedSlide(slideXml: string, analysis: SlideAnalysis, slideWidth: number, slideHeight: number) {
  const issues: string[] = [];
  const originals = new Map(analysis.shapes.map((shape) => [shape.id, shape]));
  for (const [index, match] of [...slideXml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].entries()) {
    const shape = match[0];
    const id = shapeId(shape, `shape-${index + 1}`);
    const original = originals.get(id);
    if (!original || shapeText(shape) === original.text) continue;
    const bounds = boundsOf(shape);
    if (bounds.x < 0 || bounds.y < 0 || bounds.x + bounds.width > slideWidth || bounds.y + bounds.height > slideHeight) issues.push(`shape ${id} is outside slide boundaries`);
    const allowed = Math.max(textCapacity(bounds, fontSizeOf(shape)), original.text.length);
    if (shapeText(shape).length > allowed) issues.push(`shape ${id} exceeds the template-safe text capacity`);
  }
  return issues;
}

function buildMappedSlidePackage(files: Record<string, Uint8Array>, sourceSlidePaths: string[], model: PresentationModel) {
  const dimensions = slideDimensions(files);
  const sourceSlides = sourceSlidePaths.map((path, index) => {
    const partNumber = Number(path.match(/slide(\d+)\.xml$/)?.[1]);
    const slideXml = strFromU8(files[path]);
    return {
      xml: slideXml,
      analysis: analyzeSlide(slideXml, index + 1, dimensions.width, dimensions.height),
      rels: files[`ppt/slides/_rels/slide${partNumber}.xml.rels`] ? strFromU8(files[`ppt/slides/_rels/slide${partNumber}.xml.rels`]) : "",
    };
  });
  const outputs = expandSlides(model, sourceSlides.map((source) => source.analysis));
  const sourceNotes = new Map<number, { xml: string; rels?: string }>();
  for (const source of sourceSlides) {
    const noteNumber = source.rels ? Number(notesTarget(source.rels)) : NaN;
    if (!Number.isFinite(noteNumber) || sourceNotes.has(noteNumber)) continue;
    const notePath = `ppt/notesSlides/notesSlide${noteNumber}.xml`;
    if (!files[notePath]) continue;
    const relsPath = `ppt/notesSlides/_rels/notesSlide${noteNumber}.xml.rels`;
    sourceNotes.set(noteNumber, { xml: strFromU8(files[notePath]), rels: files[relsPath] ? strFromU8(files[relsPath]) : undefined });
  }

  for (const path of Object.keys(files)) {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(path) || /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(path) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(path) || /^ppt\/notesSlides\/_rels\/notesSlide\d+\.xml\.rels$/.test(path)) delete files[path];
  }

  const packaged = outputs.map(({ slide, sourceIndex }, outputIndex) => {
    const source = sourceSlides[sourceIndex];
    const outputNumber = outputIndex + 1;
    const populated = populateSlide(source.xml, source.analysis, slide, model);
    const issues = validatePopulatedSlide(populated, source.analysis, dimensions.width, dimensions.height);
    if (issues.length) throw new Error(`Template layout validation failed on output slide ${outputNumber}: ${issues.join("; ")}.`);
    files[`ppt/slides/slide${outputNumber}.xml`] = strToU8(populated);

    let rels = source.rels;
    const sourceNoteNumber = rels ? Number(notesTarget(rels)) : NaN;
    const note = Number.isFinite(sourceNoteNumber) ? sourceNotes.get(sourceNoteNumber) : undefined;
    if (note) {
      rels = rels.replace(/Target="\.\.\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/notesSlide${outputNumber}.xml"`);
      files[`ppt/notesSlides/notesSlide${outputNumber}.xml`] = strToU8(editNotes(note.xml, slide));
      if (note.rels) files[`ppt/notesSlides/_rels/notesSlide${outputNumber}.xml.rels`] = strToU8(note.rels.replace(/Target="\.\.\/slides\/slide\d+\.xml"/, `Target="../slides/slide${outputNumber}.xml"`));
    }
    if (rels) files[`ppt/slides/_rels/slide${outputNumber}.xml.rels`] = strToU8(rels);
    return { outputNumber, hasNotes: Boolean(note) };
  });

  let presentationXml = strFromU8(files["ppt/presentation.xml"]);
  const slideIds = packaged.map(({ outputNumber }) => `<p:sldId id="${255 + outputNumber}" r:id="rIdTemplateSlide${outputNumber}"/>`).join("");
  presentationXml = presentationXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${slideIds}</p:sldIdLst>`);
  files["ppt/presentation.xml"] = strToU8(presentationXml);

  const presentationRelsPath = "ppt/_rels/presentation.xml.rels";
  let presentationRels = strFromU8(files[presentationRelsPath]);
  presentationRels = presentationRels.replace(/<Relationship\b[^>]*\/?>(?:<\/Relationship>)?/g, (relationship) => /\bType="[^"]+\/slide"/.test(relationship) ? "" : relationship);
  presentationRels = insertBeforeClosing(presentationRels, "</Relationships>", packaged.map(({ outputNumber }) => `<Relationship Id="rIdTemplateSlide${outputNumber}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${outputNumber}.xml"/>`).join(""));
  files[presentationRelsPath] = strToU8(presentationRels);

  let contentTypes = strFromU8(files["[Content_Types].xml"]);
  contentTypes = contentTypes.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "").replace(/<Override\b[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, "");
  for (const { outputNumber, hasNotes } of packaged) {
    contentTypes = addContentType(contentTypes, `/ppt/slides/slide${outputNumber}.xml`, SLIDE_CONTENT_TYPE);
    if (hasNotes) contentTypes = addContentType(contentTypes, `/ppt/notesSlides/notesSlide${outputNumber}.xml`, NOTES_CONTENT_TYPE);
  }
  files["[Content_Types].xml"] = strToU8(contentTypes);
  const appPath = "docProps/app.xml";
  if (files[appPath]) files[appPath] = strToU8(strFromU8(files[appPath]).replace(/<Slides>\d+<\/Slides>/, `<Slides>${packaged.length}</Slides>`));
}

export function fillPresentationTemplate(templateBytes: Uint8Array, model: PresentationModel) {
  const files = unzipSync(templateBytes);
  if (!files["[Content_Types].xml"] || !files["ppt/presentation.xml"]) throw new Error("The selected PPTX template is not a valid PowerPoint package.");
  const slidePaths = sortedSlidePaths(files);
  if (!slidePaths.length) throw new Error("The selected PPTX template has no slides.");
  buildMappedSlidePackage(files, slidePaths, model);
  return zipSync(files, { level: 6 });
}
