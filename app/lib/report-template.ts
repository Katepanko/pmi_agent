import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import type { ConsultingReportItem, ConsultingReportModel } from "./artifact";
import { deloitteLogoBytes, DeloitteBrand } from "./branding/deloitte.ts";

type TemplateFieldRole = "title" | "subtitle" | "audience" | "period" | "summary" | "label" | "value" | "detail" | "implication" | "recommendation" | "owner" | "deadline" | "status" | "sources";

function itemText(item: ConsultingReportItem) {
  return [
    item.value,
    item.label,
    item.detail,
    item.implication ? `Implication: ${item.implication}` : undefined,
    item.recommendation ? `Recommendation — validation required: ${item.recommendation}` : undefined,
    item.owner ? `Owner: ${item.owner}` : undefined,
    item.deadline ? `By: ${item.deadline}` : undefined,
  ].filter(Boolean).join("\n");
}

function xml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;").replace(/'/g, "&apos;");
}

function html(value: unknown) {
  return String(value ?? "").replace(/[&<>\"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character] ?? character);
}

function decodeXml(value: string) {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

function xmlText(value: string) {
  return decodeXml([...value.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join(""));
}

function excelFormula(cell: ExcelJS.Cell) {
  const value = cell.value;
  return Boolean(value && typeof value === "object" && "formula" in value);
}

function excelText(cell: ExcelJS.Cell) {
  if (cell.isMerged && cell.master.address !== cell.address) return "";
  try { return cell.text.trim(); } catch { return ""; }
}

function fieldRole(value: string): TemplateFieldRole | null {
  const text = value.toLocaleLowerCase().replace(/[_-]+/g, " ");
  if (/executive|key message|summary|takeaway/.test(text)) return "summary";
  if (/reporting period|as of|date|timing/.test(text)) return "period";
  if (/audience|prepared for|recipient/.test(text)) return "audience";
  if (/subtitle|sub title/.test(text)) return "subtitle";
  if (/title|report name/.test(text)) return "title";
  if (/management point|decision|risk|issue|action|topic|label|name/.test(text)) return "label";
  if (/status|rag|traffic light/.test(text)) return "status";
  if (/owner|accountable|responsible/.test(text)) return "owner";
  if (/deadline|due|by when|target date/.test(text)) return "deadline";
  if (/implication|impact|consequence/.test(text)) return "implication";
  if (/recommendation|next step/.test(text)) return "recommendation";
  if (/source|evidence|reference/.test(text)) return "sources";
  if (/detail|description|rationale|comment/.test(text)) return "detail";
  if (/value|metric|amount/.test(text)) return "value";
  return null;
}

function headerRole(value: string): TemplateFieldRole | null {
  const text = value.toLocaleLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  const rules: Array<[TemplateFieldRole, RegExp]> = [
    ["label", /^(?:management point|decision(?: required)?|risk|issue|action(?: required)?|topic|label|name)$/],
    ["status", /^(?:status|rag|traffic light)$/],
    ["owner", /^(?:owner|accountable|responsible)$/],
    ["deadline", /^(?:deadline|due|by when|target date)$/],
    ["implication", /^(?:implication|impact|consequence)$/],
    ["recommendation", /^(?:recommendation|next step)$/],
    ["sources", /^(?:source|sources|evidence|reference|references)$/],
    ["detail", /^(?:detail|description|rationale|comment|evidence \/ detail)$/],
    ["value", /^(?:value|metric|amount|status \/ value)$/],
  ];
  return rules.find(([, pattern]) => pattern.test(text))?.[0] ?? null;
}

function valueForRole(role: TemplateFieldRole, model: ConsultingReportModel, item?: ConsultingReportItem) {
  if (role === "title") return model.title;
  if (role === "subtitle") return model.subtitle ?? "";
  if (role === "audience") return model.audience;
  if (role === "period") return model.reportingPeriod ?? "";
  if (role === "summary") return model.executiveSummary;
  if (!item) return "";
  if (role === "label") return item.label;
  if (role === "value") return item.value ?? "";
  if (role === "detail") return item.detail ?? "Not evidenced";
  if (role === "implication") return item.implication ?? "";
  if (role === "recommendation") return item.recommendation ?? "";
  if (role === "owner") return item.owner ?? "";
  if (role === "deadline") return item.deadline ?? "";
  if (role === "status") return item.status?.toUpperCase() ?? "NOT EVIDENCED";
  return item.sourceRefs?.join(", ") ?? "";
}

/** Builds a compact, format-neutral capacity model before content generation. */
export async function analyzeReportTemplate(templateBytes: Uint8Array, fileType: string): Promise<Record<string, unknown>> {
  if (fileType === "docx") {
    const files = unzipSync(templateBytes);
    const document = files["word/document.xml"] ? strFromU8(files["word/document.xml"]) : "";
    if (!document) throw new Error("The selected Word template is not a valid DOCX package.");
    const paragraphs = [...document.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((match) => ({
      style: match[0].match(/<w:pStyle\b[^>]*w:val="([^"]+)"/)?.[1] ?? "body",
      text: xmlText(match[0]).slice(0, 160),
    })).filter((entry) => entry.text);
    const tables = [...document.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)].map((match) => {
      const rows = [...match[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)];
      const headers = rows[0] ? [...rows[0][0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cell) => xmlText(cell[0])) : [];
      return { headers, maxItems: Math.max(0, rows.length - 1) };
    });
    return { format: "docx", paragraphs, tables, headerCount: Object.keys(files).filter((path) => /^word\/header\d+\.xml$/.test(path)).length, footerCount: Object.keys(files).filter((path) => /^word\/footer\d+\.xml$/.test(path)).length };
  }
  if (fileType === "xlsx") {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBytes as unknown as ExcelJS.Buffer);
    return {
      format: "xlsx",
      worksheets: workbook.worksheets.map((sheet) => {
        let formulaCells = 0;
        const headers: Array<{ row: number; values: string[]; maxItems: number }> = [];
        sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          const values: string[] = [];
          row.eachCell({ includeEmpty: true }, (cell) => { if (excelFormula(cell)) formulaCells++; values.push(excelText(cell)); });
          if (values.filter((value) => headerRole(value)).length >= 2) headers.push({ row: rowNumber, values, maxItems: Math.max(0, sheet.rowCount - rowNumber) });
        });
        return { name: sheet.name, rows: sheet.actualRowCount, columns: sheet.actualColumnCount, formulaCells, tables: headers, frozen: (sheet.views ?? []).some((view) => view.state === "frozen") };
      }),
    };
  }
  if (fileType === "xls") return { format: "xls", strategy: "legacy-reference-only", warning: "Legacy XLS cannot be edited safely; convert it to XLSX for native template filling." };
  if (fileType === "pdf") {
    const document = await PDFDocument.load(templateBytes);
    let fields: Array<{ name: string; type: string }> = [];
    try { fields = document.getForm().getFields().map((field) => ({ name: field.getName(), type: field.constructor.name })); } catch { /* PDFs without AcroForms have no safe fill fields. */ }
    return { format: "pdf", pages: document.getPages().map((page, index) => ({ page: index + 1, ...page.getSize() })), fields, strategy: fields.length ? "form-fields" : "clean-page-reconstruction" };
  }
  if (["html", "htm"].includes(fileType)) {
    const source = new TextDecoder().decode(templateBytes);
    return {
      format: "html",
      placeholders: [...source.matchAll(/{{\s*([\w.-]+)\s*}}/g)].map((match) => match[1]),
      headings: [...source.matchAll(/<h([1-6])\b[^>]*>([\s\S]*?)<\/h\1>/gi)].map((match) => match[2].replace(/<[^>]+>/g, " ").trim()).slice(0, 30),
      tableCount: (source.match(/<table\b/gi) ?? []).length,
      hasMain: /<main\b/i.test(source),
    };
  }
  if (fileType === "csv") {
    const rows = new TextDecoder().decode(templateBytes).split(/\r?\n/).filter(Boolean);
    return { format: "csv", headers: rows[0]?.split(",") ?? [], maxItems: Math.max(0, rows.length - 1) };
  }
  return { format: fileType, strategy: "visual-reference-only" };
}

function replaceWordParagraph(paragraph: string, replacement: string) {
  let first = true;
  return paragraph.replace(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_match, attrs = "") => {
    if (!first) return `<w:t${attrs}></w:t>`;
    first = false;
    return `<w:t${attrs}>${xml(replacement)}</w:t>`;
  });
}

function fillWordTable(table: string, model: ConsultingReportModel, items: ConsultingReportItem[]) {
  const rows = [...table.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)];
  if (rows.length < 2) return table;
  const headers = [...rows[0][0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cell) => xmlText(cell[0]));
  const roles = headers.map(headerRole);
  if (roles.filter(Boolean).length < 2) return table;
  let rowIndex = 0;
  return table.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
    if (rowIndex++ === 0) return row;
    const item = items[rowIndex - 2];
    if (!item) return row.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cell) => replaceWordParagraph(cell, ""));
    let cellIndex = 0;
    return row.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cell) => {
      const role = roles[cellIndex++] ?? "detail";
      return replaceWordParagraph(cell, valueForRole(role, model, item));
    });
  });
}

export function fillWordTemplate(templateBytes: Uint8Array, model: ConsultingReportModel) {
  const files = unzipSync(templateBytes);
  const path = "word/document.xml";
  if (!files[path]) throw new Error("The selected Word template is not a valid DOCX package.");
  const items = model.sections.flatMap((section) => section.items);
  let documentXml = strFromU8(files[path]);
  const hasTitleStyle = /<w:pStyle\b[^>]*w:val="[^"]*title[^"]*"/i.test(documentXml);
  const tables: string[] = [];
  documentXml = documentXml.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, (table) => {
    const index = tables.push(fillWordTable(table, model, items)) - 1;
    return `__PMI_TEMPLATE_TABLE_${index}__`;
  });
  let firstText = true;
  let headingIndex = 0;
  let bodyIndex = 0;
  const bodyValues = [model.executiveSummary, ...model.sections.flatMap((section) => [section.keyMessage, ...section.items.map(itemText), ...(section.sourceNotes ?? [])])].filter((value): value is string => Boolean(value));
  documentXml = documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const current = xmlText(paragraph).trim();
    if (!current) return paragraph;
    const style = paragraph.match(/<w:pStyle\b[^>]*w:val="([^"]+)"/)?.[1]?.toLocaleLowerCase() ?? "";
    if ((!hasTitleStyle && firstText) || /(?:^|\W)title$/.test(style)) { firstText = false; return replaceWordParagraph(paragraph, model.title); }
    firstText = false;
    if (/subtitle/.test(style)) return replaceWordParagraph(paragraph, [model.subtitle, model.audience, model.reportingPeriod].filter(Boolean).join(" | "));
    if (/heading\s*1|heading1/.test(style)) return replaceWordParagraph(paragraph, model.sections[headingIndex++]?.title ?? current);
    if (/{{[^}]+}}|\bexample\b/i.test(current)) return replaceWordParagraph(paragraph, bodyValues[bodyIndex++] ?? "");
    if (/executive message|decisions required|risks? and issues|action required|sources? \/ limitations/i.test(current)) return paragraph;
    return paragraph;
  });
  documentXml = documentXml.replace(/__PMI_TEMPLATE_TABLE_(\d+)__/g, (_match, index: string) => tables[Number(index)] ?? "");
  files[path] = strToU8(documentXml);
  return zipSync(files, { level: 6 });
}

function copyExcelRowStyle(source: ExcelJS.Row, target: ExcelJS.Row, columnCount: number) {
  target.height = source.height;
  for (let column = 1; column <= columnCount; column++) {
    const sourceCell = source.getCell(column);
    const targetCell = target.getCell(column);
    targetCell.style = { ...sourceCell.style };
    targetCell.numFmt = sourceCell.numFmt;
  }
}

export async function fillExcelTemplate(templateBytes: Uint8Array, model: ConsultingReportModel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes as unknown as ExcelJS.Buffer);
  const items = model.sections.flatMap((section) => section.items);
  let mappedTable = false;
  workbook.eachSheet((sheet) => {
    let prominentFilled = false;
    sheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.isMerged && cell.master.address !== cell.address || excelFormula(cell)) return;
      const role = fieldRole(excelText(cell));
      const fontSize = typeof cell.font?.size === "number" ? cell.font.size : 0;
      if (!prominentFilled && (fontSize >= 18 || cell.address === "A1")) { cell.value = model.title; prominentFilled = true; return; }
      if (role && ["title", "subtitle", "summary", "audience", "period"].includes(role)) {
        const neighbor = row.getCell(cell.col + 1);
        if (!excelFormula(neighbor) && (!neighbor.value || /{{[^}]+}}|example/i.test(excelText(neighbor)))) neighbor.value = valueForRole(role, model);
      }
      if (/\bexample\b/i.test(excelText(cell)) && !role) cell.value = model.executiveSummary;
    }));

    for (let rowNumber = 1; rowNumber <= sheet.actualRowCount; rowNumber++) {
      const header = sheet.getRow(rowNumber);
      const roles: Array<TemplateFieldRole | null> = [];
      for (let column = 1; column <= sheet.actualColumnCount; column++) roles.push(headerRole(excelText(header.getCell(column))));
      if (roles.filter(Boolean).length < 2 || !roles.includes("label")) continue;
      mappedTable = true;
      const existingCapacity = Math.max(1, sheet.rowCount - rowNumber);
      const capacity = Math.max(existingCapacity, items.length);
      for (let offset = 0; offset < capacity; offset++) {
        const targetRow = sheet.getRow(rowNumber + 1 + offset);
        if (offset >= existingCapacity) copyExcelRowStyle(sheet.getRow(rowNumber + offset), targetRow, roles.length);
        const item = items[offset];
        for (let column = 1; column <= roles.length; column++) {
          const cell = targetRow.getCell(column);
          if (excelFormula(cell)) continue;
          cell.value = item && roles[column - 1] ? valueForRole(roles[column - 1] as TemplateFieldRole, model, item) : null;
        }
      }
      break;
    }
  });
  if (!mappedTable) {
    const sheet = workbook.addWorksheet("Generated PMI Data");
    sheet.addRow(["Management point", "Value", "Status", "Evidence / detail", "Implication", "Owner", "By when", "Sources"]);
    for (const item of items) sheet.addRow([item.label, item.value ?? "", item.status?.toUpperCase() ?? "NOT EVIDENCED", item.detail ?? "", item.implication ?? item.recommendation ?? "", item.owner ?? "", item.deadline ?? "", item.sourceRefs?.join(", ") ?? ""]);
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    sheet.autoFilter = { from: "A1", to: `H${Math.max(1, items.length + 1)}` };
  }
  workbook.title = model.title;
  workbook.subject = model.executiveSummary;
  workbook.creator = DeloitteBrand.name;
  return new Uint8Array(await workbook.xlsx.writeBuffer());
}

function parseCsvRow(row: string) {
  const cells: string[] = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < row.length; index++) {
    const character = row[index];
    if (character === '"' && quoted && row[index + 1] === '"') { value += '"'; index++; }
    else if (character === '"') quoted = !quoted;
    else if (character === "," && !quoted) { cells.push(value); value = ""; }
    else value += character;
  }
  cells.push(value);
  return cells;
}

export async function fillCsvTemplate(templateBytes: Uint8Array, model: ConsultingReportModel) {
  const rows = new TextDecoder().decode(templateBytes).split(/\r?\n/).filter((row) => row.trim());
  const headers = parseCsvRow(rows[0] ?? "Management point,Value,Status,Evidence / detail,Owner,By when,Sources");
  const roles = headers.map(headerRole);
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("PMI Report");
  sheet.addRow(headers);
  const items = model.sections.flatMap((section) => section.items);
  for (const item of items) sheet.addRow(roles.map((role) => role ? valueForRole(role, model, item) : ""));
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: "A1", to: `${sheet.getColumn(headers.length).letter}${Math.max(1, items.length + 1)}` };
  workbook.title = model.title;
  workbook.subject = model.executiveSummary;
  workbook.creator = DeloitteBrand.name;
  return new Uint8Array(await workbook.xlsx.writeBuffer());
}

const PDF_REPLACEMENTS: Record<string, string> = { "→": "->", "—": "-", "–": "-", "…": "...", "•": "*", "’": "'", "“": '"', "”": '"' };

function pdfSafe(value: string, font: PDFFont) {
  let output = "";
  for (const character of value) {
    const candidate = PDF_REPLACEMENTS[character] ?? character;
    try { font.encodeText(candidate); output += candidate; } catch { output += "?"; }
  }
  return output;
}

function wrap(value: string, font: PDFFont, size: number, width: number) {
  const lines: string[] = [];
  for (const paragraph of pdfSafe(value, font).split(/\r?\n/)) {
    let current = "";
    for (const word of paragraph.split(/\s+/).filter(Boolean)) {
      const candidate = current ? `${current} ${word}` : word;
      if (!current || font.widthOfTextAtSize(candidate, size) <= width) current = candidate;
      else { lines.push(current); current = word; }
    }
    if (current) lines.push(current);
  }
  return lines;
}

function drawTextBlock(page: PDFPage, value: string, options: { x: number; y: number; width: number; size: number; font: PDFFont; color: ReturnType<typeof rgb>; maxLines?: number }) {
  const lines = wrap(value, options.font, options.size, options.width).slice(0, options.maxLines ?? 20);
  let y = options.y;
  for (const line of lines) {
    page.drawText(line, { x: options.x, y, size: options.size, font: options.font, color: options.color });
    y -= options.size + 4;
  }
  return y;
}

function tryFillPdfForm(document: PDFDocument, model: ConsultingReportModel) {
  let filled = 0;
  try {
    const form = document.getForm();
    const items = model.sections.flatMap((section) => section.items);
    let itemIndex = 0;
    for (const field of form.getFields()) {
      const role = fieldRole(field.getName());
      const target = field as unknown as { setText?: (value: string) => void };
      if (!role || !target.setText) continue;
      const item = ["label", "value", "detail", "implication", "recommendation", "owner", "deadline", "status", "sources"].includes(role) ? items[itemIndex++] : undefined;
      target.setText(valueForRole(role, model, item));
      filled++;
    }
  } catch { return 0; }
  return filled;
}

export async function fillPdfTemplate(templateBytes: Uint8Array, model: ConsultingReportModel) {
  const source = await PDFDocument.load(templateBytes);
  if (!source.getPageCount()) throw new Error("The selected PDF template has no pages.");
  if (tryFillPdfForm(source, model)) {
    source.setTitle(model.title);
    source.setAuthor(DeloitteBrand.name);
    source.setSubject(model.executiveSummary);
    return { bytes: await source.save({ useObjectStreams: false }), pageCount: source.getPageCount() };
  }

  // A flat PDF has no editable semantic regions. Reconstruct clean pages at the
  // template's dimensions instead of obscuring unknown content with overlays.
  const document = await PDFDocument.create();
  document.setTitle(model.title);
  document.setAuthor(DeloitteBrand.name);
  document.setSubject(model.executiveSummary);
  const regular = await document.embedFont(StandardFonts.Helvetica);
  const bold = await document.embedFont(StandardFonts.HelveticaBold);
  const logo = await document.embedPng(deloitteLogoBytes());
  const requiredPages = Math.max(1, model.sections.length + 1);
  const sizes = source.getPages().map((page) => page.getSize());
  for (let index = 0; index < requiredPages; index++) {
    const size = sizes[Math.min(index, sizes.length - 1)];
    const page = document.addPage([size.width, size.height]);
    const { width, height } = page.getSize();
    const margin = Math.max(28, Math.min(width, height) * 0.06);
    page.drawRectangle({ x: margin, y: height - margin - 5, width: width - margin * 2, height: 5, color: rgb(0.02, 0.42, 0.22) });
    page.drawImage(logo, { x: width - margin - 92, y: height - margin - 31, width: 84, height: 16 });
    let y = height - margin - 55;
    const section = model.sections[index - 1];
    if (index === 0) {
      y = drawTextBlock(page, model.title, { x: margin + 18, y, width: width - margin * 2 - 36, size: 23, font: bold, color: rgb(0, 0, 0), maxLines: 3 });
      y -= 15;
      y = drawTextBlock(page, model.executiveSummary, { x: margin + 18, y, width: width - margin * 2 - 36, size: 13, font: bold, color: rgb(0.02, 0.42, 0.22), maxLines: 8 });
      y -= 12;
      drawTextBlock(page, [model.subtitle, model.audience, model.reportingPeriod].filter(Boolean).join(" | "), { x: margin + 18, y, width: width - margin * 2 - 36, size: 9, font: regular, color: rgb(0.35, 0.35, 0.35), maxLines: 4 });
    } else if (section) {
      y = drawTextBlock(page, section.title, { x: margin + 18, y, width: width - margin * 2 - 36, size: 18, font: bold, color: rgb(0, 0, 0), maxLines: 3 });
      if (section.keyMessage) { y -= 8; y = drawTextBlock(page, section.keyMessage, { x: margin + 18, y, width: width - margin * 2 - 36, size: 11, font: bold, color: rgb(0.02, 0.42, 0.22), maxLines: 4 }); }
      for (const item of section.items.slice(0, 8)) {
        if (y < margin + 70) break;
        y -= 13;
        y = drawTextBlock(page, itemText(item), { x: margin + 18, y, width: width - margin * 2 - 36, size: 9.5, font: regular, color: rgb(0.16, 0.16, 0.16), maxLines: 7 });
      }
    }
    page.drawText(`${index + 1} / ${requiredPages}`, { x: width - margin - 35, y: margin + 10, size: 7, font: regular, color: rgb(0.4, 0.4, 0.4) });
  }
  return { bytes: await document.save({ useObjectStreams: false }), pageCount: document.getPageCount() };
}

function generatedHtmlContent(model: ConsultingReportModel) {
  return `<main data-pmi-generated="true"><header><h1>${html(model.title)}</h1><p>${[model.subtitle, model.audience, model.reportingPeriod].filter(Boolean).map(html).join(" · ")}</p></header><section><h2>Executive summary</h2><p>${html(model.executiveSummary)}</p></section>${model.sections.map((section) => `<section><h2>${html(section.title)}</h2>${section.keyMessage ? `<p>${html(section.keyMessage)}</p>` : ""}<div class="pmi-items">${section.items.map((item) => `<article><h3>${html(item.label)}</h3>${item.value ? `<strong>${html(item.value)}</strong>` : ""}<p>${html(item.detail ?? "Not evidenced")}</p>${item.implication ? `<p><b>Implication:</b> ${html(item.implication)}</p>` : ""}${item.owner || item.deadline ? `<small>${[item.owner && `Owner: ${html(item.owner)}`, item.deadline && `By: ${html(item.deadline)}`].filter(Boolean).join(" · ")}</small>` : ""}</article>`).join("")}</div></section>`).join("")}</main>`;
}

export function fillHtmlTemplate(templateBytes: Uint8Array, model: ConsultingReportModel) {
  let source = new TextDecoder().decode(templateBytes);
  const replacements: Record<string, string> = {
    title: model.title,
    subtitle: model.subtitle ?? "",
    audience: model.audience,
    reporting_period: model.reportingPeriod ?? "",
    executive_summary: model.executiveSummary,
    sections: model.sections.map((section) => `${section.title}: ${section.items.map(itemText).join(" | ")}`).join("\n"),
  };
  let replaced = 0;
  source = source.replace(/{{\s*([\w.-]+)\s*}}/g, (match, key: string) => {
    const value = replacements[key.toLocaleLowerCase()];
    if (value === undefined) return match;
    replaced++;
    return html(value);
  });
  if (replaced) return source;
  const content = generatedHtmlContent(model);
  if (/<main\b[\s\S]*?<\/main>/i.test(source)) return source.replace(/<main\b[\s\S]*?<\/main>/i, content);
  if (/<body\b[^>]*>/i.test(source)) return source.replace(/(<body\b[^>]*>)/i, `$1${content}`);
  return `<!doctype html><html><head><meta charset="utf-8"><title>${html(model.title)}</title></head><body>${content}</body></html>`;
}
