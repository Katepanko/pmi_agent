import { PDFDocument, PDFFont, PDFPage, StandardFonts, rgb } from "pdf-lib";
import type { ConsultingReportModel } from "../artifact";
import type { SourceManifestItem } from "../pmi-prompt";
import { deloitteLogoBytes, DeloitteBrand } from "../branding/deloitte.ts";

const PAGE = { width: 612, height: 792, left: 48, right: 48, top: 54, bottom: 48 };
const COLORS = {
  ink: rgb(0, 0, 0), body: rgb(0.19, 0.19, 0.19), muted: rgb(0.46, 0.47, 0.48),
  green: rgb(0.02, 0.42, 0.22), greenPale: rgb(0.95, 0.96, 0.89), line: rgb(0.90, 0.90, 0.90), paper: rgb(0.97, 0.97, 0.97),
  amber: rgb(0.93, 0.55, 0), red: rgb(0.85, 0.16, 0.11), white: rgb(1, 1, 1),
};

const PDF_ASCII_FALLBACKS: Record<string, string> = {
  "→": "->", "←": "<-", "↔": "<->", "⇒": "=>", "⇐": "<=", "⇔": "<=>",
  "—": "-", "–": "-", "‑": "-", "−": "-", "…": "...", "•": "*",
  "“": '"', "”": '"', "„": '"', "‘": "'", "’": "'", "‚": "'", " ": " ",
};

function pdfSafe(value: string, font: PDFFont) {
  let safe = "";
  for (const character of String(value ?? "").normalize("NFC")) {
    const candidate = PDF_ASCII_FALLBACKS[character] ?? character;
    try {
      font.encodeText(candidate);
      safe += candidate;
    } catch {
      safe += "?";
    }
  }
  return safe;
}

function wrap(text: string, font: PDFFont, size: number, maxWidth: number) {
  const paragraphs = pdfSafe(String(text || ""), font).split(/\n/);
  const lines: string[] = [];
  for (const paragraph of paragraphs) {
    let current = "";
    for (const word of paragraph.split(/\s+/).filter(Boolean)) {
      const next = current ? `${current} ${word}` : word;
      if (font.widthOfTextAtSize(next, size) <= maxWidth || !current) current = next;
      else { lines.push(current); current = word; }
    }
    if (current) lines.push(current);
    if (!paragraph.trim()) lines.push("");
  }
  return lines;
}

export async function renderPdfReport(model: ConsultingReportModel, sources: SourceManifestItem[]) {
  const doc = await PDFDocument.create();
  doc.setTitle(model.title); doc.setAuthor(DeloitteBrand.name); doc.setSubject(model.executiveSummary);
  const regular = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const logo = await doc.embedPng(deloitteLogoBytes());
  let page!: PDFPage;
  let y = 0;
  const pages: PDFPage[] = [];
  const newPage = () => {
    page = doc.addPage([PAGE.width, PAGE.height]); pages.push(page); y = PAGE.height - PAGE.top - 28;
    page.drawRectangle({ x: 0, y: 0, width: 8, height: PAGE.height, color: COLORS.green });
    page.drawImage(logo, { x: PAGE.left, y: PAGE.height - 43, width: 112, height: 21 });
    page.drawText("MANAGEMENT REPORT", { x: PAGE.left, y, size: 8, font: bold, color: COLORS.green }); y -= 24;
  };
  const ensure = (height: number) => { if (y - height < PAGE.bottom) newPage(); };
  const drawLines = (text: string, options: { size?: number; font?: PDFFont; color?: ReturnType<typeof rgb>; indent?: number; gap?: number; maxWidth?: number } = {}) => {
    const size = options.size ?? 9.5; const font = options.font ?? regular; const indent = options.indent ?? 0; const gap = options.gap ?? 3;
    const lines = wrap(text, font, size, options.maxWidth ?? PAGE.width - PAGE.left - PAGE.right - indent);
    const height = lines.length * (size + gap); ensure(height + 4);
    for (const line of lines) { page.drawText(line, { x: PAGE.left + indent, y, size, font, color: options.color ?? COLORS.body }); y -= size + gap; }
    return height;
  };
  const rule = () => { ensure(12); y -= 5; page.drawLine({ start: { x: PAGE.left, y }, end: { x: PAGE.width - PAGE.right, y }, thickness: 0.6, color: COLORS.line }); y -= 10; };

  newPage();
  drawLines(model.title, { size: 24, font: bold, color: COLORS.ink, gap: 5 });
  y -= 5;
  drawLines([model.subtitle, model.audience, model.reportingPeriod].filter(Boolean).join("  |  "), { size: 9, color: COLORS.muted });
  y -= 18;
  const summaryLines = wrap(model.executiveSummary, bold, 14, PAGE.width - PAGE.left - PAGE.right - 28);
  const boxHeight = summaryLines.length * 18 + 34; ensure(boxHeight);
  page.drawRectangle({ x: PAGE.left, y: y - boxHeight + 8, width: PAGE.width - PAGE.left - PAGE.right, height: boxHeight, color: COLORS.greenPale });
  page.drawRectangle({ x: PAGE.left, y: y - boxHeight + 8, width: 5, height: boxHeight, color: COLORS.green });
  page.drawText("EXECUTIVE MESSAGE", { x: PAGE.left + 16, y: y - 8, size: 8, font: bold, color: COLORS.green });
  y -= 28;
  summaryLines.forEach((line) => { page.drawText(line, { x: PAGE.left + 16, y, size: 14, font: bold, color: COLORS.ink }); y -= 18; });
  y -= 18;

  model.sections.forEach((section, sectionIndex) => {
    // Reserve enough room for the section heading, key message, and first item so
    // conclusion-led headings never become orphans at the bottom of a page.
    ensure(150);
    if (sectionIndex) rule();
    drawLines(section.title, { size: 16, font: bold, color: COLORS.ink, gap: 4 });
    if (section.keyMessage) { y -= 2; drawLines(section.keyMessage, { size: 10.5, font: bold, color: COLORS.green, indent: 10 }); y -= 5; }
    const items = section.items.length ? section.items : [{ label: "Evidence gap", detail: "No supported detail was available for this section." }];
    for (const item of items) {
      ensure(66);
      const statusColor = item.status === "green" ? COLORS.green : item.status === "amber" ? COLORS.amber : item.status === "red" ? COLORS.red : COLORS.muted;
      page.drawRectangle({ x: PAGE.left, y: y - 3, width: 4, height: 13, color: statusColor });
      drawLines(`${item.label}${item.value ? `  |  ${item.value}` : ""}`, { size: 10, font: bold, color: COLORS.ink, indent: 12 });
      const detail = [item.detail ?? "Not evidenced", item.implication && `Implication: ${item.implication}`, item.recommendation && `Recommendation — validation required: ${item.recommendation}`, [item.owner, item.deadline].filter(Boolean).length && `Owner / timing: ${[item.owner, item.deadline].filter(Boolean).join(" | ")}`, item.sourceRefs?.length && `Sources: ${item.sourceRefs.join(", ")}`].filter(Boolean).join("  ");
      drawLines(detail, { size: 8.7, color: COLORS.body, indent: 12, maxWidth: PAGE.width - PAGE.left - PAGE.right - 12 });
      y -= 8;
    }
    if (section.sourceNotes?.length) drawLines(`Sources / limitations: ${section.sourceNotes.join(" | ")}`, { size: 7.5, color: COLORS.muted });
  });

  if (sources.length) {
    ensure(180);
    rule();
    drawLines("Evidence register", { size: 18, font: bold, color: COLORS.ink });
    drawLines("Complete source coverage and extraction limitations", { size: 9, color: COLORS.muted }); y -= 8;
    sources.forEach((source) => {
      ensure(58);
      drawLines(`${source.fileName}  |  ${source.status.toUpperCase()}`, { size: 10, font: bold, color: COLORS.ink });
      drawLines([`ID: ${source.id}`, source.excerpt ?? "Not evidenced", ...(source.warnings ?? []).map((warning) => `Warning: ${warning}`)].join("  "), { size: 8.5, color: COLORS.body });
      rule();
    });
  }

  pages.forEach((target, index) => {
    target.drawLine({ start: { x: PAGE.left, y: 31 }, end: { x: PAGE.width - PAGE.right, y: 31 }, thickness: 0.5, color: COLORS.line });
    target.drawText(`${DeloitteBrand.footer.copyright()}  •  ${DeloitteBrand.footer.confidentiality}`, { x: PAGE.left, y: 18, size: 7.5, font: regular, color: COLORS.muted });
    const count = `${index + 1} / ${pages.length}`;
    target.drawText(count, { x: PAGE.width - PAGE.right - regular.widthOfTextAtSize(count, 7.5), y: 18, size: 7.5, font: regular, color: COLORS.muted });
  });
  return { bytes: await doc.save({ useObjectStreams: false }), pageCount: pages.length };
}
