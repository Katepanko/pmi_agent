import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  PageNumber,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import type { ConsultingReportItem, ConsultingReportModel } from "../artifact";
import type { SourceManifestItem } from "../pmi-prompt";

const GREEN = "168A55";
const GREEN_PALE = "EAF7EF";
const INK = "17201B";
const BODY = "35423B";
const MUTED = "69766F";
const LINE = "D9E1DB";
const PAPER = "F5F7F5";
function cell(text: string, width: number, options: { header?: boolean; shade?: string; bold?: boolean } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.PERCENTAGE },
    shading: options.header ? { type: ShadingType.SOLID, color: INK, fill: INK } : options.shade ? { type: ShadingType.SOLID, color: options.shade, fill: options.shade } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      left: { style: BorderStyle.NONE, size: 0, color: LINE },
      right: { style: BorderStyle.NONE, size: 0, color: LINE },
    },
    children: [new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: text || "Not evidenced", bold: options.header || options.bold, color: options.header ? "FFFFFF" : BODY, size: options.header ? 18 : 17, font: "Aptos" })],
    })],
  });
}

function itemTable(items: ConsultingReportItem[]) {
  const rows = [new TableRow({ tableHeader: true, children: [
    cell("Management point", 22, { header: true }),
    cell("Status / value", 14, { header: true }),
    cell("Evidence and implication", 40, { header: true }),
    cell("Owner / timing", 24, { header: true }),
  ] })];
  (items.length ? items : [{ label: "Evidence gap", detail: "No supported detail was available for this section.", evidenceType: "gap" as const }]).forEach((item, index) => {
    const shade = index % 2 ? PAPER : "FFFFFF";
    const status = item.status && item.status !== "neutral" ? item.status.toUpperCase() : "NOT EVIDENCED";
    const evidence = [item.detail, item.implication && `Implication: ${item.implication}`, item.recommendation && `Recommendation — validation required: ${item.recommendation}`, item.sourceRefs?.length && `Sources: ${item.sourceRefs.join(", ")}`].filter(Boolean).join("\n");
    rows.push(new TableRow({ cantSplit: true, children: [
      cell(item.label, 22, { shade, bold: true }),
      cell([item.value, status].filter(Boolean).join("\n"), 14, { shade }),
      cell(evidence, 40, { shade }),
      cell([item.owner, item.deadline].filter(Boolean).join("\n"), 24, { shade }),
    ] }));
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [2200, 1400, 4000, 2400], rows });
}

export async function renderWordDocument(model: ConsultingReportModel, sources: SourceManifestItem[]) {
  const children: Array<Paragraph | Table> = [
    new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "PMI MANAGEMENT REPORT", bold: true, color: GREEN, size: 19, characterSpacing: 80, font: "Aptos" })] }),
    new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 100 }, children: [new TextRun({ text: model.title, bold: true, color: INK, size: 42, font: "Aptos Display" })] }),
    new Paragraph({ spacing: { after: 280 }, children: [new TextRun({ text: [model.subtitle, model.audience, model.reportingPeriod].filter(Boolean).join("  |  "), color: MUTED, size: 20, font: "Aptos" })] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: [new TableCell({
      shading: { type: ShadingType.SOLID, color: GREEN_PALE, fill: GREEN_PALE },
      borders: { left: { style: BorderStyle.SINGLE, size: 18, color: GREEN }, top: { style: BorderStyle.NONE, size: 0, color: GREEN_PALE }, bottom: { style: BorderStyle.NONE, size: 0, color: GREEN_PALE }, right: { style: BorderStyle.NONE, size: 0, color: GREEN_PALE } },
      margins: { top: 180, bottom: 180, left: 220, right: 220 },
      children: [
        new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "EXECUTIVE MESSAGE", bold: true, color: GREEN, size: 18, font: "Aptos" })] }),
        new Paragraph({ children: [new TextRun({ text: model.executiveSummary, bold: true, color: INK, size: 25, font: "Aptos Display" })] }),
      ],
    })] })] }),
    new Paragraph({ spacing: { after: 80 } }),
  ];

  model.sections.forEach((section) => {
    children.push(
      new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 80 }, keepNext: true, children: [new TextRun({ text: section.title, bold: true, color: INK, size: 28, font: "Aptos Display" })] }),
    );
    if (section.keyMessage) children.push(new Paragraph({ spacing: { after: 140 }, border: { left: { style: BorderStyle.SINGLE, size: 12, color: GREEN } }, indent: { left: 180 }, children: [new TextRun({ text: section.keyMessage, bold: true, color: GREEN, size: 20, font: "Aptos" })] }));
    children.push(itemTable(section.items));
    if (section.sourceNotes?.length) children.push(new Paragraph({ spacing: { before: 90, after: 120 }, children: [new TextRun({ text: `Sources / limitations: ${section.sourceNotes.join(" | ")}`, italic: true, color: MUTED, size: 15, font: "Aptos" })] }));
  });

  if (sources.length) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, spacing: { before: 300, after: 80 }, children: [new TextRun({ text: "Evidence register", bold: true, color: INK, size: 28, font: "Aptos Display" })] }));
    children.push(new Paragraph({ spacing: { after: 140 }, children: [new TextRun({ text: "Source coverage and extraction limitations", color: MUTED, size: 18, font: "Aptos" })] }));
    children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
      new TableRow({ tableHeader: true, children: [cell("Source", 30, { header: true }), cell("Status", 15, { header: true }), cell("Available evidence / warnings", 55, { header: true })] }),
      ...sources.map((source, index) => new TableRow({ cantSplit: true, children: [
        cell(`${source.id}\n${source.fileName}`, 30, { shade: index % 2 ? PAPER : "FFFFFF", bold: true }),
        cell(source.status.toUpperCase(), 15, { shade: index % 2 ? PAPER : "FFFFFF" }),
        cell([source.excerpt ?? "Not evidenced", ...(source.warnings ?? []).map((warning) => `Warning: ${warning}`)].join("\n"), 55, { shade: index % 2 ? PAPER : "FFFFFF" }),
      ] })),
    ] }));
  }

  const document = new Document({
    creator: "PMI Agent",
    title: model.title,
    subject: model.executiveSummary,
    styles: { default: { document: { run: { font: "Aptos", size: 19, color: BODY }, paragraph: { spacing: { line: 260, after: 100 } } } } },
    sections: [{
      properties: { page: { margin: { top: 900, right: 850, bottom: 850, left: 850 } } },
      headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE } }, children: [new TextRun({ text: `${model.title}  |  Confidential`, color: MUTED, size: 15, font: "Aptos" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PMI Agent  •  ", color: MUTED, size: 15 }), new TextRun({ children: [PageNumber.CURRENT], color: MUTED, size: 15 })] })] }) },
      children,
    }],
  });
  const bytes = await Packer.toBuffer(document);
  return { bytes, sectionCount: model.sections.length + (sources.length ? 1 : 0) };
}
