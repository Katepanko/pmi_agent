import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
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
import PptxGenJS from "pptxgenjs";
import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import { DELOITTE_LOGO_DATA_URI, deloitteLogoBytes, DeloitteBrand } from "../branding/deloitte.ts";
import {
  assertVisualizationDataIntegrity,
  defaultExistingContentDesignPlan,
  deriveLockedMetrics,
  type ExistingContentBlock,
  type ExistingContentDesignPlan,
  type ExistingContentFormat,
  type ExistingContentMetric,
} from "../existing-content.ts";

export type ExistingContentRenderedArtifact = {
  format: ExistingContentFormat;
  mimeType: string;
  filename: string;
  bytes: Uint8Array;
  unitCount: number;
  unitLabel: string;
  renderedTextBlocks: string[];
};

const MIME = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  html: "text/html; charset=utf-8",
} as const;

function cleanFilePart(value: string) {
  return value.replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 64) || "Exported_Content";
}

function titleFor(blocks: ExistingContentBlock[]) {
  return blocks.find((block) => block.kind === "heading")?.text ?? blocks[0]?.text.slice(0, 72) ?? "Exported Content";
}

function headingLevel(level = 1) {
  const levels = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];
  return levels[Math.max(0, Math.min(level - 1, levels.length - 1))];
}

function wordRuns(block: ExistingContentBlock, options: { bold?: boolean; color?: string; size?: number } = {}) {
  const source = block.runs?.length ? block.runs : [{ text: block.text }];
  return source.map((run) => new TextRun({
    text: run.text,
    bold: options.bold || run.bold,
    italics: run.italic || block.kind === "quote",
    font: run.code || block.kind === "code" ? "Aptos Mono" : block.kind === "heading" ? "Aptos Display" : "Aptos",
    size: options.size ?? (block.kind === "heading" ? Math.max(24, 38 - ((block.level ?? 1) - 1) * 4) : 21),
    color: options.color ?? (block.kind === "heading" ? DeloitteBrand.colors.black : "313131"),
  }));
}

function wordTable(block: ExistingContentBlock) {
  const table = block.table!;
  const width = Math.floor(100 / Math.max(1, table.headers.length));
  const cell = (text: string, header = false) => new TableCell({
    width: { size: width, type: WidthType.PERCENTAGE },
    shading: header ? { type: ShadingType.SOLID, color: "000000", fill: "000000" } : undefined,
    margins: { top: 110, bottom: 110, left: 130, right: 130 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: "D0D0CE" }, bottom: { style: BorderStyle.SINGLE, size: 2, color: "D0D0CE" },
      left: { style: BorderStyle.NONE, size: 0, color: "D0D0CE" }, right: { style: BorderStyle.NONE, size: 0, color: "D0D0CE" },
    },
    children: [new Paragraph({ children: [new TextRun({ text, bold: header, color: header ? "FFFFFF" : "313131", font: "Aptos", size: 18 })] })],
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [new TableRow({ tableHeader: true, children: table.headers.map((value) => cell(value, true)) }), ...table.rows.map((row) => new TableRow({ cantSplit: true, children: table.headers.map((_, index) => cell(row[index] ?? "")) }))],
  });
}

function wordVisualization(metrics: ExistingContentMetric[], type: "kpi_cards" | "bar_chart" | "timeline") {
  const visible = metrics.slice(0, type === "kpi_cards" ? 4 : 6);
  if (type === "kpi_cards") {
    return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: visible.map((metric) => new TableCell({
      width: { size: Math.floor(100 / visible.length), type: WidthType.PERCENTAGE },
      shading: { type: ShadingType.SOLID, color: "F1F6E4", fill: "F1F6E4" },
      margins: { top: 170, bottom: 170, left: 150, right: 150 },
      borders: { top: { style: BorderStyle.SINGLE, size: 10, color: DeloitteBrand.colors.deepGreen }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.SINGLE, size: 3, color: "FFFFFF" }, right: { style: BorderStyle.SINGLE, size: 3, color: "FFFFFF" } },
      children: [new Paragraph({ children: [new TextRun({ text: metric.value, bold: true, color: DeloitteBrand.colors.deepGreen, size: 30, font: "Aptos Display" })] }), new Paragraph({ children: [new TextRun({ text: metric.label, bold: true, color: "313131", size: 16, font: "Aptos" })] })],
    })) })] });
  }
  const maximum = Math.max(...visible.map((metric) => Math.abs(metric.numericValue)), 1);
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: visible.map((metric) => new TableRow({ children: [
    new TableCell({ width: { size: 28, type: WidthType.PERCENTAGE }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: metric.label, bold: true, size: 17, font: "Aptos" })] })] }),
    new TableCell({ width: { size: Math.max(8, Math.round(Math.abs(metric.numericValue) / maximum * 50)), type: WidthType.PERCENTAGE }, shading: { type: ShadingType.SOLID, color: "86BC25", fill: "86BC25" }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: metric.value, bold: true, size: 17, color: "000000", font: "Aptos" })] })] }),
  ] })) });
}

async function renderDocx(blocks: ExistingContentBlock[], plan: ExistingContentDesignPlan, filename: string): Promise<ExistingContentRenderedArtifact> {
  const byId = new Map(blocks.map((block) => [block.id, block]));
  const children: Array<Paragraph | Table> = [];
  const metrics = deriveLockedMetrics(blocks);
  const appendVisualizations = (pageNumber: number) => {
    for (const visualization of plan.visualizations.filter((candidate) => candidate.page === pageNumber)) {
      const supported = metrics.filter((metric) => visualization.sourceBlockIds.includes(metric.sourceBlockId));
      if (!supported.length) continue;
      children.push(new Paragraph({ spacing: { before: 220, after: 80 }, children: [new TextRun({ text: "KEY METRICS", bold: true, color: DeloitteBrand.colors.deepGreen, size: 17, characterSpacing: 80, font: "Aptos" })] }));
      children.push(wordVisualization(supported, visualization.type));
    }
  };
  let currentPage = 1;
  for (const placement of [...plan.placements].sort((left, right) => left.page - right.page)) {
    if (placement.page > currentPage) {
      appendVisualizations(currentPage);
      children.push(new Paragraph({ pageBreakBefore: true }));
      currentPage = placement.page;
    }
    for (const id of placement.blockIds) {
      const block = byId.get(id)!;
      if (block.kind === "table" && block.table) { children.push(wordTable(block)); continue; }
      if (placement.component === "title") {
        children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: [new TableCell({ shading: { type: ShadingType.SOLID, color: "000000", fill: "000000" }, margins: { top: 500, bottom: 500, left: 380, right: 380 }, borders: { top: { style: BorderStyle.NONE, size: 0, color: "000000" }, bottom: { style: BorderStyle.SINGLE, size: 18, color: "86BC25" }, left: { style: BorderStyle.NONE, size: 0, color: "000000" }, right: { style: BorderStyle.NONE, size: 0, color: "000000" } }, children: [new Paragraph({ children: wordRuns(block, { bold: true, color: "FFFFFF", size: 44 }) })] })] })] }));
      } else if (placement.component === "executive_message" || placement.component === "callout") {
        children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: [new TableCell({ shading: { type: ShadingType.SOLID, color: "F1F6E4", fill: "F1F6E4" }, margins: { top: 210, bottom: 210, left: 250, right: 250 }, borders: { left: { style: BorderStyle.SINGLE, size: 18, color: DeloitteBrand.colors.deepGreen }, top: { style: BorderStyle.NONE, size: 0, color: "F1F6E4" }, bottom: { style: BorderStyle.NONE, size: 0, color: "F1F6E4" }, right: { style: BorderStyle.NONE, size: 0, color: "F1F6E4" } }, children: [new Paragraph({ children: wordRuns(block, { bold: placement.emphasis === "high", size: 23 }) })] })] })] }));
      } else {
        children.push(new Paragraph({
          heading: block.kind === "heading" ? headingLevel(block.level) : undefined,
          bullet: block.kind === "bullet" ? { level: 0 } : undefined,
          numbering: block.kind === "numbered" ? { reference: "existing-content-numbering", level: 0 } : undefined,
          indent: block.kind === "quote" ? { left: 480 } : undefined,
          border: block.kind === "heading" ? { bottom: { style: BorderStyle.SINGLE, size: 5, color: block.level === 2 ? "86BC25" : "D0D0CE" } } : undefined,
          spacing: { before: block.kind === "heading" ? 280 : 30, after: block.kind === "heading" ? 120 : 150, line: 276 },
          keepNext: block.kind === "heading",
          children: wordRuns(block, { bold: block.kind === "heading" }),
        }));
      }
      children.push(new Paragraph({ spacing: { after: 35 } }));
    }
  }
  appendVisualizations(currentPage);
  const document = new Document({
    creator: DeloitteBrand.name,
    title: titleFor(blocks),
    numbering: { config: [{ reference: "existing-content-numbering", levels: [{ level: 0, format: "decimal", text: "%1.", alignment: AlignmentType.LEFT }] }] },
    sections: [{
      properties: { page: { margin: { top: 900, right: 850, bottom: 850, left: 850 } } },
      headers: { default: new Header({ children: [new Paragraph({ children: [new ImageRun({ data: deloitteLogoBytes(), transformation: { width: 132, height: 25 }, type: "png" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], color: DeloitteBrand.colors.coolGray, size: 16 })] })] }) },
      children,
    }],
  });
  const bytes = await Packer.toBuffer(document);
  return { format: "docx", mimeType: MIME.docx, filename, bytes: new Uint8Array(bytes), unitCount: Math.max(...plan.placements.map((placement) => placement.page), 1), unitLabel: "pages", renderedTextBlocks: blocks.map((block) => block.text) };
}

async function renderPptx(blocks: ExistingContentBlock[], plan: ExistingContentDesignPlan, filename: string): Promise<ExistingContentRenderedArtifact> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = DeloitteBrand.name;
  pptx.title = titleFor(blocks);
  pptx.subject = "Verbatim export of existing conversation content";
  pptx.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos" };
  const byId = new Map(blocks.map((block) => [block.id, block]));
  const metrics = deriveLockedMetrics(blocks);
  const pages = [...new Set(plan.placements.map((placement) => placement.page))].sort((a, b) => a - b);
  pages.forEach((pageNumber, slideIndex) => {
    const placements = plan.placements.filter((placement) => placement.page === pageNumber);
    const slide = pptx.addSlide();
    const cover = placements.length === 1 && placements[0].component === "title";
    slide.background = { color: cover ? "000000" : "FFFFFF" };
    slide.addShape("rect", { x: 0, y: 0, w: cover ? 13.333 : 0.16, h: cover ? 0.14 : 7.5, line: { transparency: 100 }, fill: { color: DeloitteBrand.colors.deepGreen } });
    slide.addImage({ data: DELOITTE_LOGO_DATA_URI, x: cover ? 10.9 : 11.1, y: 0.34, w: 1.45, h: 0.27 });
    let y = cover ? 1.3 : 0.72;
    for (const placement of placements) {
      for (const id of placement.blockIds) {
        const block = byId.get(id)!;
        const heading = block.kind === "heading";
        const x = placement.column === "right" ? 6.85 : placement.column === "left" ? 0.72 : 0.72;
        const w = placement.column === "full" ? 11.85 : 5.75;
        if (block.table) {
          slide.addTable([block.table.headers, ...block.table.rows], { x, y, w, h: Math.min(3.3, 0.42 * (block.table.rows.length + 1)), border: { color: "D0D0CE", pt: 0.6 }, fill: "FFFFFF", color: "313131", fontFace: "Aptos", fontSize: 10, bold: false, margin: 0.08, rowH: 0.4, autoFit: false });
          y += Math.min(3.3, 0.42 * (block.table.rows.length + 1)) + 0.2;
          continue;
        }
        const lines = Math.max(1, Math.ceil(block.text.length / (heading ? 65 : 105)));
        const height = placement.component === "executive_message" ? Math.min(1.55, Math.max(0.72, lines * 0.3 + 0.3)) : Math.min(cover ? 2.2 : heading ? 1.05 : 1.4, Math.max(heading ? 0.48 : 0.34, lines * (heading ? 0.38 : 0.24)));
        if (placement.component === "executive_message" || placement.component === "callout") {
          slide.addShape("roundRect", { x, y: y - 0.08, w, h: height + 0.18, rectRadius: 0.04, line: { color: "D6E6C3", pt: 0.7 }, fill: { color: "F1F6E4" } });
          slide.addShape("rect", { x, y: y - 0.08, w: 0.08, h: height + 0.18, line: { transparency: 100 }, fill: { color: DeloitteBrand.colors.deepGreen } });
        }
        const prefix = block.kind === "bullet" ? "•  " : block.kind === "numbered" ? `${blocks.filter((candidate) => candidate.kind === "numbered" && Number(candidate.id.slice(6)) <= Number(block.id.slice(6))).length}.  ` : "";
        const richText = block.runs?.length && !prefix ? block.runs.map((run) => ({ text: run.text, options: { bold: run.bold, italic: run.italic, fontFace: run.code ? "Aptos Mono" : undefined } })) : `${prefix}${block.text}`;
        slide.addText(richText, {
          x: x + (placement.component === "executive_message" || placement.component === "callout" ? 0.25 : 0), y, w: w - (placement.component === "executive_message" || placement.component === "callout" ? 0.4 : 0), h: height,
          fontFace: block.kind === "code" ? "Aptos Mono" : heading ? "Aptos Display" : "Aptos",
          fontSize: cover ? 32 : heading ? Math.max(18, 27 - ((block.level ?? 1) - 1) * 2) : placement.emphasis === "high" ? 16 : 13.5,
          bold: heading || placement.emphasis === "high", italic: block.kind === "quote", color: cover ? "FFFFFF" : heading ? "000000" : "313131", margin: 0, valign: "top", fit: "shrink",
        });
        if (heading && !cover) slide.addShape("line", { x, y: y + height + 0.03, w: Math.min(w, 2.1), h: 0, line: { color: DeloitteBrand.colors.green, width: 2.2 } });
        y += height + (heading ? 0.25 : 0.18);
      }
    }
    for (const visualization of plan.visualizations.filter((candidate) => candidate.page === pageNumber)) {
      const supported = metrics.filter((metric) => visualization.sourceBlockIds.includes(metric.sourceBlockId)).slice(0, 5);
      if (!supported.length) continue;
      const visualY = Math.min(y + 0.1, 5.25);
      if (visualization.type === "kpi_cards") {
        const cardWidth = Math.min(2.55, 11.8 / supported.length);
        supported.forEach((metric, index) => {
          const cardX = 0.72 + index * (cardWidth + 0.12);
          slide.addShape("roundRect", { x: cardX, y: visualY, w: cardWidth, h: 1.05, rectRadius: 0.04, line: { color: "D0D0CE", pt: 0.6 }, fill: { color: "F1F6E4" } });
          slide.addText(metric.value, { x: cardX + 0.16, y: visualY + 0.15, w: cardWidth - 0.32, h: 0.38, fontFace: "Aptos Display", fontSize: 23, bold: true, color: DeloitteBrand.colors.deepGreen, margin: 0, fit: "shrink" });
          slide.addText(metric.label, { x: cardX + 0.16, y: visualY + 0.65, w: cardWidth - 0.32, h: 0.2, fontFace: "Aptos", fontSize: 9, bold: true, color: "313131", margin: 0, fit: "shrink" });
        });
      } else {
        const maximum = Math.max(...supported.map((metric) => Math.abs(metric.numericValue)), 1);
        supported.forEach((metric, index) => {
          const rowY = visualY + index * 0.34;
          slide.addText(metric.label, { x: 0.74, y: rowY, w: 2.45, h: 0.18, fontFace: "Aptos", fontSize: 8.5, bold: true, color: "313131", margin: 0, fit: "shrink" });
          slide.addShape("rect", { x: 3.2, y: rowY, w: Math.max(0.25, Math.abs(metric.numericValue) / maximum * 6.8), h: 0.2, line: { transparency: 100 }, fill: { color: DeloitteBrand.colors.green } });
          slide.addText(metric.value, { x: 10.15, y: rowY, w: 1.3, h: 0.18, fontFace: "Aptos", fontSize: 8.5, bold: true, color: "000000", margin: 0, fit: "shrink" });
        });
      }
    }
    slide.addText(`${slideIndex + 1} / ${pages.length}`, { x: 11.72, y: 7.12, w: 0.8, h: 0.16, fontFace: "Aptos", fontSize: 7, color: cover ? "8FA096" : DeloitteBrand.colors.coolGray, align: "right", margin: 0 });
  });
  const result = await pptx.write({ outputType: "uint8array", compression: true });
  const bytes = result instanceof Uint8Array ? result : result instanceof ArrayBuffer ? new Uint8Array(result) : result instanceof Blob ? new Uint8Array(await result.arrayBuffer()) : null;
  if (!bytes) throw new Error("The PowerPoint renderer returned an unsupported output type.");
  return { format: "pptx", mimeType: MIME.pptx, filename, bytes, unitCount: pages.length, unitLabel: "slides", renderedTextBlocks: blocks.map((block) => block.text) };
}

function wrapPdf(text: string, font: PDFFont, size: number, width: number) {
  const output: string[] = [];
  for (const paragraph of text.split("\n")) {
    let line = "";
    for (const word of paragraph.split(/\s+/).filter(Boolean)) {
      const candidate = line ? `${line} ${word}` : word;
      if (!line || font.widthOfTextAtSize(candidate, size) <= width) line = candidate;
      else { output.push(line); line = word; }
    }
    if (line) output.push(line);
    if (!paragraph.trim()) output.push("");
  }
  return output;
}

async function renderPdf(blocks: ExistingContentBlock[], plan: ExistingContentDesignPlan, filename: string): Promise<ExistingContentRenderedArtifact> {
  const document = await PDFDocument.create();
  document.setTitle(titleFor(blocks));
  document.setAuthor(DeloitteBrand.name);
  const regular = await document.embedFont(StandardFonts.Helvetica);
  const bold = await document.embedFont(StandardFonts.HelveticaBold);
  const logo = await document.embedPng(deloitteLogoBytes());
  const PAGE = { width: 612, height: 792, left: 52, right: 52, top: 54, bottom: 48 };
  let page: PDFPage = document.addPage([PAGE.width, PAGE.height]);
  let y = PAGE.height - PAGE.top;
  let pageCount = 1;
  const decoratePage = () => {
    y = PAGE.height - PAGE.top;
    page.drawRectangle({ x: 0, y: 0, width: 7, height: PAGE.height, color: rgb(0.02, 0.42, 0.22) });
    page.drawImage(logo, { x: PAGE.left, y: PAGE.height - 37, width: 105, height: 20 });
    y -= 20;
  };
  const newPage = () => {
    page = document.addPage([PAGE.width, PAGE.height]);
    pageCount += 1;
    decoratePage();
  };
  decoratePage();
  const byId = new Map(blocks.map((block) => [block.id, block]));
  let plannedPage = 1;
  for (const placement of [...plan.placements].sort((left, right) => left.page - right.page)) {
    while (placement.page > plannedPage) { newPage(); plannedPage += 1; }
    for (const id of placement.blockIds) {
      const block = byId.get(id)!;
      if (block.table) {
        const columns = Math.max(1, block.table.headers.length);
        const columnWidth = (PAGE.width - PAGE.left - PAGE.right) / columns;
        const rows = [block.table.headers, ...block.table.rows];
        const required = rows.length * 28 + 18;
        if (y - required < PAGE.bottom) newPage();
        rows.forEach((row, rowIndex) => {
          row.forEach((cell, columnIndex) => {
            page.drawRectangle({ x: PAGE.left + columnIndex * columnWidth, y: y - 22, width: columnWidth, height: 26, color: rowIndex === 0 ? rgb(0, 0, 0) : rowIndex % 2 ? rgb(0.97, 0.97, 0.97) : rgb(1, 1, 1), borderColor: rgb(0.82, 0.82, 0.81), borderWidth: 0.4 });
            const lines = wrapPdf(cell, rowIndex === 0 ? bold : regular, 8.5, columnWidth - 12).slice(0, 2);
            lines.forEach((line, lineIndex) => page.drawText(line, { x: PAGE.left + columnIndex * columnWidth + 6, y: y - 8 - lineIndex * 10, size: 8.5, font: rowIndex === 0 ? bold : regular, color: rowIndex === 0 ? rgb(1, 1, 1) : rgb(0.19, 0.19, 0.19) }));
          });
          y -= 26;
        });
        y -= 14;
        continue;
      }
      const heading = block.kind === "heading";
      const cover = placement.component === "title";
      const font = heading || placement.emphasis === "high" ? bold : regular;
      const size = cover ? 25 : heading ? Math.max(13, 22 - ((block.level ?? 1) - 1) * 2) : placement.emphasis === "high" ? 12.5 : 10.5;
      let lines: string[];
      try {
        lines = wrapPdf(block.text, font, size, PAGE.width - PAGE.left - PAGE.right - (block.kind === "quote" ? 20 : 0));
      } catch {
        throw new Error(`PDF export cannot encode one or more characters in locked block ${block.id}; the text was not altered.`);
      }
      const lineHeight = size + 4;
      const needed = lines.length * lineHeight + (heading ? 24 : 18);
      if (y - needed < PAGE.bottom) newPage();
      if (cover) {
        page.drawRectangle({ x: PAGE.left, y: y - needed - 25, width: PAGE.width - PAGE.left - PAGE.right, height: needed + 55, color: rgb(0, 0, 0) });
        page.drawRectangle({ x: PAGE.left, y: y - needed - 25, width: PAGE.width - PAGE.left - PAGE.right, height: 7, color: rgb(0.53, 0.74, 0.15) });
        y -= 22;
      } else if (placement.component === "executive_message" || placement.component === "callout") {
        page.drawRectangle({ x: PAGE.left, y: y - needed + 5, width: PAGE.width - PAGE.left - PAGE.right, height: needed + 8, color: rgb(0.95, 0.97, 0.90) });
        page.drawRectangle({ x: PAGE.left, y: y - needed + 5, width: 5, height: needed + 8, color: rgb(0.02, 0.42, 0.22) });
      }
      const indent = block.kind === "quote" || block.kind === "bullet" || block.kind === "numbered" || placement.component === "executive_message" || placement.component === "callout" ? 16 : 0;
      if (block.kind === "bullet") page.drawText("•", { x: PAGE.left, y, size, font: regular, color: rgb(0.19, 0.19, 0.19) });
      if (block.kind === "numbered") {
        const count = blocks.filter((candidate) => candidate.kind === "numbered" && Number(candidate.id.slice(6)) <= Number(block.id.slice(6))).length;
        page.drawText(`${count}.`, { x: PAGE.left, y, size, font: regular, color: rgb(0.19, 0.19, 0.19) });
      }
      for (const line of lines) {
        page.drawText(line, { x: PAGE.left + indent, y, size, font, color: cover ? rgb(1, 1, 1) : heading ? rgb(0, 0, 0) : rgb(0.19, 0.19, 0.19) });
        y -= lineHeight;
      }
      if (heading && !cover) page.drawLine({ start: { x: PAGE.left, y: y - 2 }, end: { x: PAGE.left + 130, y: y - 2 }, thickness: 2, color: rgb(0.53, 0.74, 0.15) });
      y -= heading ? 16 : 10;
    }
  }
  const metrics = deriveLockedMetrics(blocks);
  for (const visualization of plan.visualizations) {
    const supported = metrics.filter((metric) => visualization.sourceBlockIds.includes(metric.sourceBlockId)).slice(0, 6);
    if (!supported.length) continue;
    const needed = visualization.type === "kpi_cards" ? 82 : supported.length * 25 + 30;
    if (y - needed < PAGE.bottom) newPage();
    page.drawText("KEY METRICS", { x: PAGE.left, y, size: 8, font: bold, color: rgb(0.02, 0.42, 0.22) });
    y -= 20;
    if (visualization.type === "kpi_cards") {
      const width = (PAGE.width - PAGE.left - PAGE.right - (supported.length - 1) * 8) / supported.length;
      supported.forEach((metric, index) => {
        const x = PAGE.left + index * (width + 8);
        page.drawRectangle({ x, y: y - 52, width, height: 58, color: rgb(0.95, 0.97, 0.90), borderColor: rgb(0.82, 0.82, 0.81), borderWidth: 0.5 });
        page.drawText(metric.value, { x: x + 9, y: y - 17, size: 16, font: bold, color: rgb(0.02, 0.42, 0.22) });
        wrapPdf(metric.label, bold, 7.5, width - 18).slice(0, 2).forEach((line, lineIndex) => page.drawText(line, { x: x + 9, y: y - 35 - lineIndex * 9, size: 7.5, font: bold, color: rgb(0.19, 0.19, 0.19) }));
      });
      y -= 72;
    } else {
      const maximum = Math.max(...supported.map((metric) => Math.abs(metric.numericValue)), 1);
      supported.forEach((metric) => {
        page.drawText(metric.label, { x: PAGE.left, y, size: 8.5, font: bold, color: rgb(0.19, 0.19, 0.19) });
        page.drawRectangle({ x: PAGE.left + 150, y: y - 2, width: Math.max(12, Math.abs(metric.numericValue) / maximum * 250), height: 10, color: rgb(0.53, 0.74, 0.15) });
        page.drawText(metric.value, { x: PAGE.width - PAGE.right - 65, y, size: 8.5, font: bold, color: rgb(0, 0, 0) });
        y -= 24;
      });
    }
  }
  const bytes = await document.save({ useObjectStreams: false });
  return { format: "pdf", mimeType: MIME.pdf, filename, bytes, unitCount: pageCount, unitLabel: "pages", renderedTextBlocks: blocks.map((block) => block.text) };
}

function escapeHtml(value: string) {
  return value.replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]!);
}

function htmlRuns(block: ExistingContentBlock) {
  if (!block.runs?.length) return escapeHtml(block.text).replace(/\n/g, "<br>");
  return block.runs.map((run) => {
    const value = escapeHtml(run.text).replace(/\n/g, "<br>");
    if (run.bold) return `<strong>${value}</strong>`;
    if (run.italic) return `<em>${value}</em>`;
    if (run.code) return `<code>${value}</code>`;
    return value;
  }).join("");
}

function htmlBlock(block: ExistingContentBlock, component: string) {
  const text = htmlRuns(block);
  const attributes = `data-block-id="${escapeHtml(block.id)}" class="component ${component}"`;
  if (block.table) return `<div ${attributes}><table><thead><tr>${block.table.headers.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead><tbody>${block.table.rows.map((row) => `<tr>${block.table!.headers.map((_, index) => `<td>${escapeHtml(row[index] ?? "")}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
  if (block.kind === "heading") return `<h${Math.max(1, Math.min(6, block.level ?? 1))} ${attributes}>${text}</h${Math.max(1, Math.min(6, block.level ?? 1))}>`;
  if (block.kind === "bullet") return `<div ${attributes}><div class="list bullet"><span aria-hidden="true">•</span><p>${text}</p></div></div>`;
  if (block.kind === "numbered") return `<div ${attributes}><div class="list numbered"><p>${text}</p></div></div>`;
  if (block.kind === "quote") return `<blockquote ${attributes}>${text}</blockquote>`;
  if (block.kind === "code") return `<pre ${attributes}><code>${text}</code></pre>`;
  return `<div ${attributes}><p>${text}</p></div>`;
}

function htmlVisualization(metrics: ExistingContentMetric[], type: "kpi_cards" | "bar_chart" | "timeline") {
  if (type === "kpi_cards") return `<div class="kpis">${metrics.slice(0, 4).map((metric) => `<article class="kpi"><strong>${escapeHtml(metric.value)}</strong><span>${escapeHtml(metric.label)}</span></article>`).join("")}</div>`;
  const maximum = Math.max(...metrics.map((metric) => Math.abs(metric.numericValue)), 1);
  return `<div class="chart">${metrics.slice(0, 6).map((metric) => `<div class="bar-row"><b>${escapeHtml(metric.label)}</b><i style="--bar:${Math.max(4, Math.abs(metric.numericValue) / maximum * 100)}%"></i><strong>${escapeHtml(metric.value)}</strong></div>`).join("")}</div>`;
}

function renderHtml(blocks: ExistingContentBlock[], plan: ExistingContentDesignPlan, filename: string): ExistingContentRenderedArtifact {
  const title = escapeHtml(titleFor(blocks));
  const byId = new Map(blocks.map((block) => [block.id, block]));
  const metrics = deriveLockedMetrics(blocks);
  const pages = [...new Set(plan.placements.map((placement) => placement.page))].sort((a, b) => a - b);
  const content = pages.map((page) => {
    const placements = plan.placements.filter((placement) => placement.page === page);
    const narrative = placements.map((placement) => `<div class="placement ${placement.column}">${placement.blockIds.map((id) => htmlBlock(byId.get(id)!, placement.component)).join("\n")}</div>`).join("\n");
    const visuals = plan.visualizations.filter((visualization) => visualization.page === page).map((visualization) => htmlVisualization(metrics.filter((metric) => visualization.sourceBlockIds.includes(metric.sourceBlockId)), visualization.type)).join("\n");
    return `<section class="report-page" data-page="${page}">${narrative}${visuals}</section>`;
  }).join("\n");
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${title}</title><style>:root{--green:#046a38;--lime:#86bc25;--ink:#000;--body:#313131;--line:#d0d0ce;--pale:#f1f6e4}*{box-sizing:border-box}body{margin:0;color:var(--body);font:17px/1.55 Aptos,"Segoe UI",sans-serif;background:#eceeed}.shell{max-width:1040px;margin:24px auto;background:white;box-shadow:0 18px 54px #0002;border-top:8px solid var(--green)}.brandbar{height:78px;padding:24px 7%;border-bottom:1px solid var(--line)}.brand{width:142px;height:auto}.report-page{padding:46px 7% 64px;min-height:580px;border-bottom:12px solid #eceeed;display:grid;grid-template-columns:1fr 1fr;gap:0 26px;align-content:start}.placement.full,.kpis,.chart{grid-column:1/-1}.placement.left{grid-column:1}.placement.right{grid-column:2}.component{margin-bottom:20px}h1,h2,h3,h4,h5,h6{color:var(--ink);font-family:"Aptos Display",Aptos,sans-serif;line-height:1.14;margin:1.1em 0 .42em}h1.title{font-size:48px;background:#000;color:white;padding:54px 44px;border-bottom:7px solid var(--lime);margin:0 0 30px}h2,h3{padding-bottom:9px;border-bottom:3px solid var(--lime)}h2{font-size:31px}p{margin:0;white-space:normal}.executive_message,.callout{background:var(--pale);border-left:7px solid var(--green);padding:24px 26px;font-size:20px;font-weight:650}.list{display:flex;gap:12px;margin:0 0 8px;padding-left:10px}.list p{margin:0}.numbered{display:list-item;list-style:decimal;margin-left:32px;padding-left:4px}blockquote{border-left:5px solid var(--green);margin:1em 0;padding:14px 20px;background:var(--pale);font-style:italic}pre{padding:16px;background:#f1f1f1;white-space:pre-wrap;overflow-wrap:anywhere}table{width:100%;border-collapse:collapse;margin:10px 0 22px;font-size:14px}th{background:#000;color:white;text-align:left;padding:11px 12px}td{padding:10px 12px;border-bottom:1px solid var(--line)}tr:nth-child(even) td{background:#f7f7f7}.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:24px 0}.kpi{padding:18px;border:1px solid var(--line);border-top:5px solid var(--lime);background:var(--pale)}.kpi strong{display:block;color:var(--green);font:700 28px/1 "Aptos Display",sans-serif}.kpi span{display:block;margin-top:9px;font-size:12px;font-weight:800;text-transform:uppercase}.chart{margin:26px 0;padding:22px;border:1px solid var(--line)}.bar-row{display:grid;grid-template-columns:180px 1fr 90px;gap:12px;align-items:center;margin:10px 0}.bar-row b{font-size:13px}.bar-row i{display:block;width:var(--bar);height:18px;background:var(--lime)}.bar-row strong{text-align:right}@media(max-width:700px){.shell{margin:0}.report-page{display:block;padding:32px 6%}.kpis{grid-template-columns:1fr 1fr}.bar-row{grid-template-columns:120px 1fr 70px}}@media print{body{background:white}.shell{max-width:none;margin:0;box-shadow:none}.report-page{break-after:page;border:0}}</style></head><body><main class="shell"><header class="brandbar"><img class="brand" alt="Deloitte" src="${DELOITTE_LOGO_DATA_URI}"></header>${content}</main></body></html>`;
  return { format: "html", mimeType: MIME.html, filename, bytes: new TextEncoder().encode(html), unitCount: pages.length, unitLabel: "sections", renderedTextBlocks: blocks.map((block) => block.text) };
}

export async function renderExistingContent(input: { format: ExistingContentFormat; blocks: ExistingContentBlock[]; version: number; plan?: ExistingContentDesignPlan }): Promise<ExistingContentRenderedArtifact> {
  const filename = `${cleanFilePart(titleFor(input.blocks))}${input.version > 1 ? `_v${input.version}` : ""}.${input.format}`;
  const plan = input.plan ?? defaultExistingContentDesignPlan(input.blocks);
  assertVisualizationDataIntegrity(input.blocks, plan);
  if (input.format === "docx") return renderDocx(input.blocks, plan, filename);
  if (input.format === "pptx") return renderPptx(input.blocks, plan, filename);
  if (input.format === "pdf") return renderPdf(input.blocks, plan, filename);
  return renderHtml(input.blocks, plan, filename);
}
