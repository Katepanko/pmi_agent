import { strFromU8, unzipSync } from "fflate";
import type { ExtractedDocument } from "./types";

const supported = new Set(["xlsx", "xls", "csv", "pptx", "docx", "pdf", "html", "htm", "png", "jpg", "jpeg"]);

function decodeXml(value: string) {
  return value
    .replace(/<w:tab\/?\s*>/g, "\t")
    .replace(/<a:br\/?\s*>/g, "\n")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)));
}

function textNodes(xml: string, tagPattern = "(?:w:t|a:t|t)") {
  const pattern = new RegExp(`<${tagPattern}[^>]*>([\\s\\S]*?)<\\/(?:w:t|a:t|t)>`, "g");
  return [...xml.matchAll(pattern)].map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")));
}

function fileExtension(fileName: string) {
  return fileName.split(".").pop()?.toLowerCase() ?? "";
}

function baseDocument(file: File, type: string): ExtractedDocument {
  return {
    fileId: crypto.randomUUID(),
    fileName: file.name,
    fileType: type,
    mediaType: file.type || "application/octet-stream",
    sizeBytes: file.size,
    status: "extracted",
    metadata: {},
    rawText: "",
    structuredElements: [],
    tables: [],
    sheets: [],
    slides: [],
    pages: [],
    images: [],
    extractionWarnings: [],
  };
}

function columnName(index: number) {
  let result = "";
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

function parseCsv(text: string) {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === '"' && quoted && text[index + 1] === '"') {
      field += '"';
      index += 1;
    } else if (character === '"') {
      quoted = !quoted;
    } else if (character === "," && !quoted) {
      row.push(field); field = "";
    } else if ((character === "\n" || character === "\r") && !quoted) {
      if (character === "\r" && text[index + 1] === "\n") index += 1;
      row.push(field); field = "";
      if (row.some((cell) => cell.length)) rows.push(row);
      row = [];
    } else {
      field += character;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function extractCsv(file: File, text: string) {
  const document = baseDocument(file, "csv");
  const rows = parseCsv(text);
  const headers = rows[0] ?? [];
  document.metadata = { rowCount: Math.max(0, rows.length - 1), columnCount: headers.length, headers };
  document.tables = [{ name: file.name, headers, rowCount: Math.max(0, rows.length - 1), rows: rows.slice(1) }];
  document.structuredElements = rows.slice(1).map((values, index) => {
    const record = Object.fromEntries(headers.map((header, column) => [header || columnName(column + 1), values[column] ?? ""]));
    return {
      ordinal: index,
      location: `${file.name} → row ${index + 2}`,
      kind: "record",
      content: headers.map((header, column) => `${header || columnName(column + 1)}: ${values[column] ?? ""}`).join(" | "),
      structured: { row: index + 2, values: record },
      confidence: "high",
    };
  });
  document.rawText = document.structuredElements.map((segment) => segment.content).join("\n");
  return document;
}

function extractHtml(file: File, html: string) {
  const document = baseDocument(file, fileExtension(file.name));
  const withoutNoise = html.replace(/<script[\s\S]*?<\/script>/gi, " ").replace(/<style[\s\S]*?<\/style>/gi, " ");
  const blocks = [...withoutNoise.matchAll(/<(h[1-6]|p|li|tr)[^>]*>([\s\S]*?)<\/\1>/gi)]
    .map((match) => decodeXml(match[2].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()))
    .filter(Boolean);
  document.structuredElements = blocks.map((content, ordinal) => ({
    ordinal,
    location: `${file.name} → HTML block ${ordinal + 1}`,
    kind: "text",
    content,
    confidence: "high",
  }));
  document.rawText = blocks.join("\n");
  document.metadata = { blockCount: blocks.length };
  if (!blocks.length) document.extractionWarnings.push("No semantic HTML blocks were found.");
  return document;
}

function unzipOffice(bytes: Uint8Array) {
  return unzipSync(bytes);
}

function extractDocx(file: File, bytes: Uint8Array) {
  const document = baseDocument(file, "docx");
  const archive = unzipOffice(bytes);
  const xmlBytes = archive["word/document.xml"];
  if (!xmlBytes) throw new Error("word/document.xml is missing.");
  const xml = strFromU8(xmlBytes);
  const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
    .map((match) => textNodes(match[0], "w:t").join(""))
    .filter((text) => text.trim());
  document.structuredElements = paragraphs.map((content, ordinal) => ({
    ordinal,
    location: `${file.name} → paragraph ${ordinal + 1}`,
    kind: "text",
    content,
    structured: transcriptFields(content),
    confidence: "high",
  }));
  document.rawText = paragraphs.join("\n");
  document.metadata = { paragraphCount: paragraphs.length, embeddedMediaCount: Object.keys(archive).filter((name) => name.startsWith("word/media/")).length };
  document.images = Object.keys(archive).filter((name) => name.startsWith("word/media/")).map((name) => ({ location: name, requiresVision: true }));
  if (document.images.length) {
    document.status = "partial";
    document.extractionWarnings.push(`${document.images.length} embedded image(s) require multimodal inspection.`);
  }
  return document;
}

function transcriptFields(content: string) {
  const timestamp = content.match(/\b(?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?\b/)?.[0];
  const speaker = content.match(/^([^:]{2,60}):\s+/)?.[1];
  const lower = content.toLowerCase();
  const classification = [
    ["decision", /\b(decided|decision|agreed|approved)\b/],
    ["action", /\b(action|will|follow up|owner)\b/],
    ["risk", /\b(risk|may|could|threat)\b/],
    ["issue", /\b(issue|failed|blocked|overdue)\b/],
    ["dependency", /\b(depend|dependency|prerequisite)\b/],
  ].find(([, pattern]) => (pattern as RegExp).test(lower))?.[0];
  return { speaker, timestamp, classification };
}

function extractPptx(file: File, bytes: Uint8Array) {
  const document = baseDocument(file, "pptx");
  const archive = unzipOffice(bytes);
  const slideNames = Object.keys(archive)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));

  slideNames.forEach((name, index) => {
    const slideNumber = index + 1;
    const text = textNodes(strFromU8(archive[name]), "a:t").join("\n").trim();
    const notesName = `ppt/notesSlides/notesSlide${slideNumber}.xml`;
    const notes = archive[notesName] ? textNodes(strFromU8(archive[notesName]), "a:t").join("\n").trim() : "";
    const content = [text, notes ? `Speaker notes: ${notes}` : ""].filter(Boolean).join("\n");
    document.slides.push({ slideNumber, title: text.split("\n")[0] ?? "", text, notes });
    document.structuredElements.push({ ordinal: index, location: `${file.name} → slide ${slideNumber}`, kind: "text", content, confidence: "high" });
  });
  document.rawText = document.structuredElements.map((segment) => segment.content).join("\n\n");
  document.metadata = { slideCount: slideNames.length };
  const media = Object.keys(archive).filter((name) => name.startsWith("ppt/media/"));
  document.images = media.map((name) => ({ location: name, requiresVision: true }));
  if (media.length) {
    document.status = "partial";
    document.extractionWarnings.push(`${media.length} slide visual asset(s) require multimodal inspection for complete visual context.`);
  }
  return document;
}

function extractXlsx(file: File, bytes: Uint8Array) {
  const document = baseDocument(file, "xlsx");
  const archive = unzipOffice(bytes);
  const workbook = archive["xl/workbook.xml"] ? strFromU8(archive["xl/workbook.xml"]) : "";
  const sharedXml = archive["xl/sharedStrings.xml"] ? strFromU8(archive["xl/sharedStrings.xml"]) : "";
  const sharedStrings = [...sharedXml.matchAll(/<si\b[\s\S]*?<\/si>/g)].map((match) => textNodes(match[0], "t").join(""));
  const sheetNames = [...workbook.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*\/>/g)].map((match) => decodeXml(match[1]));
  const worksheets = Object.keys(archive).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name)).sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
  let ordinal = 0;

  worksheets.forEach((name, sheetIndex) => {
    const sheetName = sheetNames[sheetIndex] ?? `Sheet ${sheetIndex + 1}`;
    const xml = strFromU8(archive[name]);
    const rowRecords: Array<Record<string, unknown>> = [];
    for (const rowMatch of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
      const rowNumber = Number(rowMatch[1].match(/\br="(\d+)"/)?.[1] ?? rowRecords.length + 1);
      const cells: Record<string, unknown> = {};
      const formulas: Record<string, string> = {};
      for (const cellMatch of rowMatch[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attrs = cellMatch[1];
        const cellRef = attrs.match(/\br="([^"]+)"/)?.[1] ?? `${columnName(Object.keys(cells).length + 1)}${rowNumber}`;
        const type = attrs.match(/\bt="([^"]+)"/)?.[1];
        const raw = cellMatch[2].match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? "";
        const formula = cellMatch[2].match(/<f[^>]*>([\s\S]*?)<\/f>/)?.[1];
        let value: string | number = decodeXml(raw);
        if (type === "s") value = sharedStrings[Number(raw)] ?? raw;
        else if (type === "inlineStr") value = textNodes(cellMatch[2], "t").join("");
        else if (raw !== "" && Number.isFinite(Number(raw))) value = Number(raw);
        cells[cellRef] = value;
        if (formula) formulas[cellRef] = decodeXml(formula);
      }
      if (!Object.keys(cells).length) continue;
      rowRecords.push({ row: rowNumber, cells, formulas });
      document.structuredElements.push({
        ordinal: ordinal++,
        location: `${file.name} → sheet "${sheetName}", row ${rowNumber}`,
        kind: "record",
        content: Object.entries(cells).map(([cell, value]) => `${cell}: ${String(value)}`).join(" | "),
        structured: { sheet: sheetName, row: rowNumber, cells, formulas },
        confidence: "high",
      });
    }
    document.sheets.push({ name: sheetName, rowCount: rowRecords.length, rows: rowRecords });
  });
  document.rawText = document.structuredElements.map((segment) => `${segment.location}\n${segment.content}`).join("\n");
  document.metadata = { sheetCount: worksheets.length, sheetNames, sharedStringCount: sharedStrings.length };
  return document;
}

function extractImage(file: File, bytes: Uint8Array) {
  const type = fileExtension(file.name);
  const document = baseDocument(file, type);
  const dimensions = imageDimensions(type, bytes);
  document.status = "partial";
  document.metadata = { ...dimensions };
  document.images = [{ location: `${file.name} → image`, requiresVision: true, ...dimensions }];
  document.structuredElements = [{ ordinal: 0, location: `${file.name} → image`, kind: "image", content: "Image queued for multimodal inspection.", structured: dimensions, confidence: "low" }];
  document.extractionWarnings.push("Text, charts, and dashboard meaning require multimodal inspection before coverage can be complete.");
  return document;
}

function imageDimensions(type: string, bytes: Uint8Array) {
  if (type === "png" && bytes.length >= 24) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (["jpg", "jpeg"].includes(type)) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) break;
      const marker = bytes[offset + 1];
      const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
        return { height: (bytes[offset + 5] << 8) + bytes[offset + 6], width: (bytes[offset + 7] << 8) + bytes[offset + 8] };
      }
      offset += 2 + length;
    }
  }
  return {};
}

function extractPdf(file: File, bytes: Uint8Array) {
  const document = baseDocument(file, "pdf");
  const binary = new TextDecoder("latin1").decode(bytes);
  const pageCount = Math.max(1, [...binary.matchAll(/\/Type\s*\/Page\b/g)].length);
  const text = [...binary.matchAll(/\(((?:\\.|[^\\)])*)\)\s*Tj/g)]
    .map((match) => match[1].replace(/\\([()\\])/g, "$1"))
    .join(" ");
  document.status = text.trim() ? "partial" : "partial";
  document.metadata = { pageCount };
  document.pages = Array.from({ length: pageCount }, (_, index) => ({ pageNumber: index + 1 }));
  document.rawText = text;
  if (text.trim()) document.structuredElements.push({ ordinal: 0, location: `${file.name} → PDF text layer`, kind: "text", content: text, confidence: "low" });
  document.extractionWarnings.push("PDF extraction used a conservative text-layer pass; complex layout, compressed text, tables, and images require the dedicated PDF renderer/parser stage.");
  return document;
}

export async function extractFile(file: File, requestedFileId?: string): Promise<ExtractedDocument> {
  const type = fileExtension(file.name);
  if (!supported.has(type)) {
    const document = baseDocument(file, type || "unknown");
    document.status = "failed";
    document.extractionWarnings.push(`Unsupported file type: ${type || "unknown"}.`);
    if (requestedFileId) document.fileId = requestedFileId;
    return document;
  }

  try {
    let extracted: ExtractedDocument;
    if (type === "csv") extracted = extractCsv(file, await file.text());
    else if (type === "html" || type === "htm") extracted = extractHtml(file, await file.text());
    else {
    const bytes = new Uint8Array(await file.arrayBuffer());
      if (type === "docx") extracted = extractDocx(file, bytes);
      else if (type === "pptx") extracted = extractPptx(file, bytes);
      else if (type === "xlsx") extracted = extractXlsx(file, bytes);
      else if (type === "png" || type === "jpg" || type === "jpeg") extracted = extractImage(file, bytes);
      else if (type === "pdf") extracted = extractPdf(file, bytes);
      else {
        extracted = baseDocument(file, type);
        extracted.status = "partial";
        extracted.extractionWarnings.push("Legacy .xls requires the server-side spreadsheet compatibility adapter; the source remains pending and cannot be marked covered.");
      }
    }
    if (requestedFileId) extracted.fileId = requestedFileId;
    return extracted;
  } catch (error) {
    const document = baseDocument(file, type);
    if (requestedFileId) document.fileId = requestedFileId;
    document.status = "failed";
    document.extractionWarnings.push(error instanceof Error ? error.message : "Extraction failed.");
    return document;
  }
}

export async function extractFiles(files: File[]) {
  return Promise.all(files.map(extractFile));
}

export function createCoverageCheck(documents: ExtractedDocument[]) {
  const sourcesConsidered = documents.map((document) => ({
    fileId: document.fileId,
    fileName: document.fileName,
    status: document.status === "extracted" ? "used" : document.status === "partial" ? "pending_validation" : "not_used",
    relevantSections: document.structuredElements.map((segment) => segment.location).slice(0, 25),
    warnings: document.extractionWarnings,
  }));
  const sourcesNotUsed = sourcesConsidered.filter((source) => source.status === "not_used");
  return {
    requestId: crypto.randomUUID(),
    sourcesConsidered,
    sourcesNotUsed,
    coverageComplete: documents.length > 0 && documents.every((document) => document.status === "extracted"),
  };
}
