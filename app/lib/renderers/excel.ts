import ExcelJS from "exceljs";
import type { ConsultingReportModel, ConsultingReportSection } from "../artifact";
import type { SourceManifestItem } from "../pmi-prompt";
import { DELOITTE_LOGO_DATA_URI, DeloitteBrand } from "../branding/deloitte.ts";

const COLORS = {
  ink: `FF${DeloitteBrand.colors.black}`,
  body: "FF313131",
  muted: `FF${DeloitteBrand.colors.coolGray}`,
  white: `FF${DeloitteBrand.colors.white}`,
  green: `FF${DeloitteBrand.colors.brightGreen}`,
  greenDark: `FF${DeloitteBrand.colors.deepGreen}`,
  greenPale: `FF${DeloitteBrand.colors.paleGreen}`,
  amber: `FF${DeloitteBrand.colors.amber}`,
  amberPale: "FFFFF4DE",
  red: `FF${DeloitteBrand.colors.red}`,
  redPale: "FFFDECEC",
  paper: "FFF7F7F7",
  line: `FF${DeloitteBrand.colors.lightGray}`,
};

function safeSheetName(value: string, used: Set<string>, index: number) {
  const base = value.replace(/[\\/*?:[\]]/g, " ").replace(/\s+/g, " ").trim().slice(0, 26) || `Section ${index}`;
  let candidate = `${String(index).padStart(2, "0")} ${base}`.slice(0, 31);
  let suffix = 2;
  while (used.has(candidate)) candidate = `${candidate.slice(0, 28)} ${suffix++}`.slice(0, 31);
  used.add(candidate);
  return candidate;
}

function statusLabel(value?: string) {
  return value === "green" ? "ON TRACK" : value === "amber" ? "AT RISK" : value === "red" ? "CRITICAL" : "NOT EVIDENCED";
}

function styleTitle(sheet: ExcelJS.Worksheet, model: ConsultingReportModel, subtitle: string, logoId: number) {
  sheet.mergeCells("A1:F2");
  const title = sheet.getCell("A1");
  title.value = model.title;
  title.font = { name: "Aptos Display", size: 22, bold: true, color: { argb: COLORS.white } };
  title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.ink } };
  title.alignment = { vertical: "middle", horizontal: "left" };
  sheet.getRow(1).height = 27;
  sheet.getRow(2).height = 27;
  sheet.mergeCells("G1:H2");
  const logoPanel = sheet.getCell("G1");
  logoPanel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.white } };
  sheet.addImage(logoId, { tl: { col: 6.18, row: 0.45 }, ext: { width: 142, height: 27 }, editAs: "oneCell" });
  sheet.mergeCells("A3:H3");
  const meta = sheet.getCell("A3");
  meta.value = `${subtitle}  |  ${model.audience}${model.reportingPeriod ? `  |  ${model.reportingPeriod}` : ""}`;
  meta.font = { name: "Aptos", size: 10, bold: true, color: { argb: COLORS.greenDark } };
  meta.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.greenPale } };
  meta.alignment = { vertical: "middle" };
  sheet.getRow(3).height = 22;
}

function styleTableHeader(row: ExcelJS.Row) {
  row.height = 27;
  row.eachCell((cell) => {
    cell.font = { name: "Aptos", size: 9, bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.ink } };
    cell.alignment = { vertical: "middle", wrapText: true };
  });
}

function addStatusConditionalFormatting(sheet: ExcelJS.Worksheet, range: string) {
  sheet.addConditionalFormatting({
    ref: range,
    rules: [
      { type: "containsText", operator: "containsText", text: "ON TRACK", style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: COLORS.greenPale }, fgColor: { argb: COLORS.greenPale } }, font: { color: { argb: COLORS.greenDark }, bold: true } } },
      { type: "containsText", operator: "containsText", text: "AT RISK", style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: COLORS.amberPale }, fgColor: { argb: COLORS.amberPale } }, font: { color: { argb: COLORS.amber }, bold: true } } },
      { type: "containsText", operator: "containsText", text: "CRITICAL", style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: COLORS.redPale }, fgColor: { argb: COLORS.redPale } }, font: { color: { argb: COLORS.red }, bold: true } } },
    ],
  });
}

function configureSheet(sheet: ExcelJS.Worksheet) {
  sheet.views = [{ showGridLines: false, state: "frozen", ySplit: 6 }];
  sheet.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.35, right: 0.35, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
  sheet.headerFooter.oddFooter = `&L${DeloitteBrand.footer.copyright()}&C${DeloitteBrand.footer.confidentiality}&RPage &P of &N`;
  const widths = [26, 17, 16, 46, 46, 22, 18, 24];
  widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
}

function addExecutiveSummary(workbook: ExcelJS.Workbook, model: ConsultingReportModel, sections: ConsultingReportSection[], logoId: number) {
  const sheet = workbook.addWorksheet("01 Executive Summary", { properties: { tabColor: { argb: COLORS.green } } });
  configureSheet(sheet);
  sheet.views = [{ showGridLines: false, state: "frozen", ySplit: 11 }];
  styleTitle(sheet, model, "EXECUTIVE MANAGEMENT VIEW", logoId);
  sheet.mergeCells("A5:H6");
  const message = sheet.getCell("A5");
  message.value = model.executiveSummary;
  message.font = { name: "Aptos Display", size: 15, bold: true, color: { argb: COLORS.ink } };
  message.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.white } };
  message.border = { left: { style: "thick", color: { argb: COLORS.green } } };
  message.alignment = { vertical: "middle", wrapText: true };
  sheet.getRow(5).height = 26;
  sheet.getRow(6).height = 26;

  const allItems = sections.flatMap((section) => section.items);
  const counts = {
    "ON TRACK": allItems.filter((item) => item.status === "green").length,
    "AT RISK": allItems.filter((item) => item.status === "amber").length,
    CRITICAL: allItems.filter((item) => item.status === "red").length,
    "NOT EVIDENCED": allItems.filter((item) => !item.status || item.status === "neutral").length,
  };
  const cards = [
    ["A8:B8", "A9:B10", "ON TRACK", counts["ON TRACK"], COLORS.greenDark, COLORS.greenPale],
    ["C8:D8", "C9:D10", "AT RISK", counts["AT RISK"], COLORS.amber, COLORS.amberPale],
    ["E8:F8", "E9:F10", "CRITICAL", counts.CRITICAL, COLORS.red, COLORS.redPale],
    ["G8:H8", "G9:H10", "NOT EVIDENCED", counts["NOT EVIDENCED"], COLORS.muted, COLORS.paper],
  ] as const;
  for (const [labelRange, valueRange, label, value, color, fill] of cards) {
    sheet.mergeCells(labelRange);
    sheet.mergeCells(valueRange);
    const labelCell = sheet.getCell(labelRange.split(":")[0]);
    labelCell.value = label;
    labelCell.font = { name: "Aptos", size: 9, bold: true, color: { argb: color } };
    labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    labelCell.alignment = { horizontal: "center", vertical: "middle" };
    const valueCell = sheet.getCell(valueRange.split(":")[0]);
    valueCell.value = value;
    valueCell.font = { name: "Aptos Display", size: 24, bold: true, color: { argb: color } };
    valueCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    valueCell.alignment = { horizontal: "center", vertical: "middle" };
  }

  sheet.mergeCells("A12:H12");
  sheet.getCell("A12").value = "MANAGEMENT MESSAGES";
  sheet.getCell("A12").font = { name: "Aptos", size: 10, bold: true, color: { argb: COLORS.white } };
  sheet.getCell("A12").fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.ink } };
  const messages = sections.flatMap((section) => section.items.map((item) => ({ section: section.name, item }))).slice(0, 8);
  messages.forEach(({ section, item }, index) => {
    const row = 13 + index;
    sheet.mergeCells(`A${row}:B${row}`);
    sheet.mergeCells(`C${row}:H${row}`);
    sheet.getCell(`A${row}`).value = section;
    sheet.getCell(`A${row}`).font = { name: "Aptos", size: 9, bold: true, color: { argb: COLORS.ink } };
    sheet.getCell(`C${row}`).value = `${item.label}${item.implication ? ` - ${item.implication}` : item.detail ? ` - ${item.detail}` : ""}`;
    sheet.getCell(`C${row}`).font = { name: "Aptos", size: 9, color: { argb: COLORS.body } };
    sheet.getCell(`C${row}`).alignment = { wrapText: true, vertical: "middle" };
    sheet.getRow(row).height = 34;
    for (let column = 1; column <= 8; column++) sheet.getCell(row, column).border = { bottom: { style: "hair", color: { argb: COLORS.line } } };
  });
  return sheet;
}

function addSectionSheet(workbook: ExcelJS.Workbook, model: ConsultingReportModel, section: ConsultingReportSection, name: string, logoId: number) {
  const sheet = workbook.addWorksheet(name, { properties: { tabColor: { argb: section.type === "risks" ? COLORS.red : section.type === "decisions" ? COLORS.amber : COLORS.green } } });
  configureSheet(sheet);
  styleTitle(sheet, model, section.name.toUpperCase(), logoId);
  sheet.mergeCells("A5:H5");
  sheet.getCell("A5").value = section.keyMessage ?? section.title;
  sheet.getCell("A5").font = { name: "Aptos Display", size: 13, bold: true, color: { argb: COLORS.ink } };
  sheet.getCell("A5").alignment = { vertical: "middle", wrapText: true };
  sheet.getRow(5).height = 34;
  const headers = ["Management point", "Value", "Status", "Evidence / detail", "Management implication", "Owner", "By when", "Evidence / sources"];
  sheet.getRow(6).values = headers;
  styleTableHeader(sheet.getRow(6));

  const items = section.items.length ? section.items : [{ label: "Not evidenced", detail: "No supported detail was available for this section.", status: "neutral" as const, evidenceType: "gap" as const }];
  items.forEach((item, index) => {
    const rowIndex = 7 + index;
    const row = sheet.getRow(rowIndex);
    row.values = [
      item.label,
      item.value ?? "",
      statusLabel(item.status),
      item.detail ?? "Not evidenced",
      item.implication ?? item.recommendation ?? "",
      item.owner ?? "",
      item.deadline ?? "",
      `${(item.evidenceType ?? "inference").toUpperCase()}${item.sourceRefs?.length ? ` | ${item.sourceRefs.join(", ")}` : ""}`,
    ];
    row.height = 42;
    row.eachCell((cell, column) => {
      cell.font = { name: "Aptos", size: column === 1 ? 9.5 : 9, bold: column === 1, color: { argb: COLORS.body } };
      cell.alignment = { vertical: "middle", wrapText: true, horizontal: [2, 3, 6, 7].includes(column) ? "center" : "left" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: index % 2 ? COLORS.paper : COLORS.white } };
      cell.border = { bottom: { style: "hair", color: { argb: COLORS.line } } };
    });
  });
  const lastRow = 6 + items.length;
  sheet.autoFilter = { from: "A6", to: `H${lastRow}` };
  sheet.getColumn(3).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    if (rowNumber >= 7) cell.dataValidation = { type: "list", allowBlank: true, formulae: ['"ON TRACK,AT RISK,CRITICAL,NOT EVIDENCED"'] };
  });
  addStatusConditionalFormatting(sheet, `C7:C${Math.max(lastRow, 7)}`);
  if (section.sourceNotes?.length) {
    const noteRow = lastRow + 2;
    sheet.mergeCells(`A${noteRow}:H${noteRow + 1}`);
    sheet.getCell(`A${noteRow}`).value = `Sources / limitations: ${section.sourceNotes.join(" | ")}`;
    sheet.getCell(`A${noteRow}`).font = { name: "Aptos", size: 8, italic: true, color: { argb: COLORS.muted } };
    sheet.getCell(`A${noteRow}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.paper } };
    sheet.getCell(`A${noteRow}`).alignment = { vertical: "middle", wrapText: true };
  }
  return sheet;
}

function addSourceRegister(workbook: ExcelJS.Workbook, model: ConsultingReportModel, sources: SourceManifestItem[], index: number, logoId: number) {
  if (!sources.length) return;
  const sheet = workbook.addWorksheet(`${String(index).padStart(2, "0")} Source Register`.slice(0, 31), { properties: { tabColor: { argb: COLORS.muted } } });
  configureSheet(sheet);
  styleTitle(sheet, model, "SOURCE COVERAGE AND LIMITATIONS", logoId);
  sheet.mergeCells("A5:H5");
  sheet.getCell("A5").value = "Every applicable source is listed; incomplete extraction remains explicit.";
  sheet.getCell("A5").font = { name: "Aptos Display", size: 13, bold: true, color: { argb: COLORS.ink } };
  sheet.getRow(6).values = ["Source ID", "File", "Status", "Warnings", "Available evidence", "", "", ""];
  styleTableHeader(sheet.getRow(6));
  sources.forEach((source, sourceIndex) => {
    const row = sheet.getRow(7 + sourceIndex);
    row.values = [source.id, source.fileName, source.status.toUpperCase(), (source.warnings ?? []).join(" | "), source.excerpt ?? "Not evidenced"];
    row.height = 44;
    row.eachCell((cell) => { cell.font = { name: "Aptos", size: 9, color: { argb: COLORS.body } }; cell.alignment = { vertical: "middle", wrapText: true }; cell.border = { bottom: { style: "hair", color: { argb: COLORS.line } } }; });
  });
  sheet.getColumn(1).width = 24;
  sheet.getColumn(2).width = 35;
  sheet.getColumn(3).width = 15;
  sheet.getColumn(4).width = 45;
  sheet.getColumn(5).width = 70;
  sheet.autoFilter = { from: "A6", to: `E${6 + sources.length}` };
}

export async function renderExcelWorkbook(model: ConsultingReportModel, sources: SourceManifestItem[]) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = DeloitteBrand.name;
  workbook.company = DeloitteBrand.name;
  workbook.title = model.title;
  workbook.subject = model.executiveSummary;
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  const logoId = workbook.addImage({ base64: DELOITTE_LOGO_DATA_URI, extension: "png" });
  addExecutiveSummary(workbook, model, model.sections, logoId);
  const used = new Set<string>(["01 Executive Summary"]);
  model.sections.forEach((section, index) => addSectionSheet(workbook, model, section, safeSheetName(section.name, used, index + 2), logoId));
  addSourceRegister(workbook, model, sources, model.sections.length + 2, logoId);
  const buffer = await workbook.xlsx.writeBuffer();
  return { bytes: new Uint8Array(buffer), sheetCount: workbook.worksheets.length };
}
