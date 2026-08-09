import { strToU8, zipSync } from "fflate";

export type ReportFormat = "pptx" | "pdf" | "xlsx" | "docx" | "html";

export type ReportExportInput = {
  title: string;
  audience: string;
  content: string;
  sources: Array<{ name: string; status: string }>;
};

const MIME: Record<ReportFormat, string> = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  html: "text/html;charset=utf-8",
};

function xml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function slug(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60) || "pmi-report";
}

function lines(input: ReportExportInput) {
  const content = input.content === "Illustrative report preview"
    ? "Integration momentum is broadly intact, but two cross-functional dependencies require SteerCo intervention.\nMost planned workstreams are advancing, with pressure concentrated at the IT–Finance boundary.\nSynergy initiatives remain active; timing assumptions require validation.\nDecision required: confirm the accountable executive for the cutover dependency.\nDecision required: reconcile the latest synergy timing with the baseline tracker.\nRecommendation: publish one reconciled cutover baseline with a single accountable owner."
    : input.content;
  return content.split(/\n+/).map((line) => line.replace(/^#{1,6}\s*/, "").replace(/^[-*]\s*/, "").trim()).filter(Boolean);
}

function zipBlob(files: Record<string, string>, type: string) {
  const archive = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, strToU8(value)])), { level: 6 });
  return new Blob([archive], { type });
}

function pptx(input: ReportExportInput) {
  const reportLines = lines(input);
  const slideBodies = [
    ["Executive summary", reportLines.slice(0, 3)],
    ["Decisions and management focus", reportLines.slice(3, 6)],
    ["Source coverage", input.sources.length ? input.sources.map((source) => `${source.name} — ${source.status}`) : ["No sources attached"]],
  ] as const;
  const shape = (text: string, x: number, y: number, cx: number, cy: number, size: number, color = "17231B", bold = false) => `<p:sp><p:nvSpPr><p:cNvPr id="${Math.round(x + y + size)}" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="${size * 100}" b="${bold ? 1 : 0}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${xml(text)}</a:t></a:r><a:endParaRPr lang="en-US" sz="${size * 100}"/></a:p></p:txBody></p:sp>`;
  const slideXml = (title: string, body: readonly string[], index: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F5F7F4"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Accent"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="320000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="18A862"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>${shape(title, 700000, 600000, 10500000, 800000, 30, "101713", true)}${body.slice(0, 6).map((item, itemIndex) => shape(`•  ${item}`, 900000, 1800000 + itemIndex * 720000, 9800000, 600000, 18, "334139")).join("")}${shape(`PMI Agent  ·  ${input.audience}  ·  ${index + 1}`, 700000, 6300000, 10500000, 300000, 9, "68756D")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const slides = [[input.title, [input.audience, "Decision-ready integration reporting"]] as const, ...slideBodies];
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("")}</Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("")}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${slides.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join("")}</Relationships>`,
  };
  slides.forEach(([title, body], index) => { files[`ppt/slides/slide${index + 1}.xml`] = slideXml(title, body, index); });
  return zipBlob(files, MIME.pptx);
}

function docx(input: ReportExportInput) {
  const paragraphs = lines(input).map((value) => `<w:p><w:pPr><w:spacing w:after="160"/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>${xml(value)}</w:t></w:r></w:p>`).join("");
  return zipBlob({
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="101713"/><w:sz w:val="40"/></w:rPr><w:t>${xml(input.title)}</w:t></w:r></w:p><w:p><w:r><w:rPr><w:color w:val="18A862"/><w:b/></w:rPr><w:t>PMI Agent · ${xml(input.audience)}</w:t></w:r></w:p>${paragraphs}<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Source coverage</w:t></w:r></w:p>${input.sources.map((source) => `<w:p><w:r><w:t>• ${xml(source.name)} — ${xml(source.status)}</w:t></w:r></w:p>`).join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>`,
  }, MIME.docx);
}

function xlsx(input: ReportExportInput) {
  const rows = [["PMI Agent report", input.title], ["Audience", input.audience], ["Generated", new Date().toISOString()], ["", ""], ["Report content", ""], ...lines(input).map((value, index) => [`${index + 1}`, value]), ["", ""], ["Source", "Status"], ...input.sources.map((source) => [source.name, source.status])];
  const cell = (value: string, column: string, row: number, header = false) => `<c r="${column}${row}" t="inlineStr"${header ? ' s="1"' : ""}><is><t>${xml(value)}</t></is></c>`;
  const sheet = rows.map((row, index) => `<row r="${index + 1}">${cell(row[0], "A", index + 1, [1, 5, rows.length - input.sources.length].includes(index + 1))}${cell(row[1], "B", index + 1, [1, 5, rows.length - input.sources.length].includes(index + 1))}</row>`).join("");
  return zipBlob({
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Management Report" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF101713"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="80" customWidth="1"/></cols><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>${sheet}</sheetData><autoFilter ref="A${rows.length - input.sources.length}:B${rows.length}"/></worksheet>`,
  }, MIME.xlsx);
}

function pdf(input: ReportExportInput) {
  const safe = (value: string) => value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)").replace(/[^\x20-\x7E]/g, "-");
  const wrapped = [input.title, `Audience: ${input.audience}`, "", ...lines(input), "", "Source coverage", ...input.sources.map((source) => `${source.name} - ${source.status}`)].flatMap((value) => value.match(/.{1,82}(?:\s|$)/g)?.map((part) => part.trim()) ?? [value]);
  const commands = wrapped.slice(0, 44).map((value, index) => `BT /F1 ${index === 0 ? 20 : 10} Tf 54 ${780 - index * 16} Td (${safe(value)}) Tj ET`).join("\n");
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${commands.length} >>\nstream\n${commands}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let output = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((object, index) => { offsets.push(output.length); output += `${index + 1} 0 obj\n${object}\nendobj\n`; });
  const xref = output.length;
  output += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n `).join("\n")}\ntrailer << /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return new Blob([output], { type: MIME.pdf });
}

function html(input: ReportExportInput) {
  const cards = lines(input).map((value, index) => `<article><span>${String(index + 1).padStart(2, "0")}</span><p>${xml(value)}</p></article>`).join("");
  const value = `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>${xml(input.title)}</title><style>body{margin:0;background:#f5f7f4;color:#101713;font:16px/1.5 Arial,sans-serif}header{padding:64px max(6vw,32px);background:#101713;color:white;border-bottom:8px solid #18a862}header small{color:#7de0a7;text-transform:uppercase;letter-spacing:.14em}h1{max-width:900px;font-size:clamp(36px,6vw,72px);line-height:1;margin:.3em 0}main{max-width:1100px;margin:48px auto;padding:0 24px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px}article{background:white;border:1px solid #dce3dd;border-radius:14px;padding:24px}article span{color:#18a862;font-weight:bold}.sources{margin-top:32px;padding:24px;background:#e5f7ed;border-radius:14px}.sources li{margin:.5em 0}</style><header><small>PMI Agent · ${xml(input.audience)}</small><h1>${xml(input.title)}</h1><p>Interactive management-report dashboard</p></header><main><div class="grid">${cards}</div><section class="sources"><h2>Source coverage</h2><ul>${input.sources.map((source) => `<li>${xml(source.name)} — ${xml(source.status)}</li>`).join("") || "<li>No sources attached</li>"}</ul></section></main></html>`;
  return new Blob([value], { type: MIME.html });
}

export function createReportArtifact(format: ReportFormat, input: ReportExportInput) {
  const blob = format === "pptx" ? pptx(input) : format === "docx" ? docx(input) : format === "xlsx" ? xlsx(input) : format === "pdf" ? pdf(input) : html(input);
  return { blob, fileName: `${slug(input.title)}.${format}` };
}
