import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, unzipSync } from "fflate";
import PptxGenJS from "pptxgenjs";
import ExcelJS from "exceljs";
import { buildPresentationPlanningPrompt, renderPresentation } from "../app/lib/presentation.ts";
import { analyzePresentationTemplate, fillPresentationTemplate } from "../app/lib/presentation-template.ts";
import { hasUnresolvedTemplateDirective, messageMentionsTemplate, resolveTemplateReference, templateOutputFormat } from "../app/lib/template.ts";
import { analyzeReportTemplate, fillCsvTemplate, fillExcelTemplate, fillHtmlTemplate, fillPdfTemplate, fillWordTemplate } from "../app/lib/report-template.ts";
import { renderExcelWorkbook } from "../app/lib/renderers/excel.ts";
import { renderPdfReport } from "../app/lib/renderers/pdf.ts";
import { renderWordDocument } from "../app/lib/renderers/word.ts";
import { PDFDocument } from "pdf-lib";

const templateSource = {
  id: "template-1",
  fileName: "Nike_adidas_PMI_Executive_Status.pptx",
  fileType: "pptx",
  status: "extracted",
  excerpt: "Nike and Adidas example report",
  metadata: { slideCount: 2 },
};

async function decisionTemplate() {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.addText("Example decisions should unlock delivery", { x: 0.65, y: 0.35, w: 11.8, h: 0.5, fontSize: 28, bold: true, margin: 0 });
  slide.addText("Decisions required", { x: 0.65, y: 0.95, w: 4, h: 0.3, fontSize: 14, margin: 0 });
  [0.65, 4.55, 8.45].forEach((x, index) => {
    slide.addShape("roundRect", { x, y: 1.42, w: 3.55, h: 1.28, fill: { color: "F5F7F5" }, line: { color: "D9DEDA" } });
    slide.addText("Decision required", { x: x + 0.18, y: 1.55, w: 3.1, h: 0.2, fontSize: 9, bold: true, margin: 0 });
    slide.addText(`Example decision ${index + 1}`, { x: x + 0.18, y: 1.82, w: 3.1, h: 0.3, fontSize: 15, bold: true, margin: 0 });
    slide.addText("Example rationale and implication that must be replaced.", { x: x + 0.18, y: 2.18, w: 3.1, h: 0.34, fontSize: 9, margin: 0 });
  });
  slide.addTable([
    ["Action required", "Owner", "By when", "Status"],
    ["Example action 1", "Example", "Example", "Green"],
    ["Example action 2", "Example", "Example", "Amber"],
    ["Example action 3", "Example", "Example", "Red"],
    ["Example action 4", "Example", "Example", "Green"],
    ["Example action 5", "Example", "Example", "Amber"],
  ], { x: 0.65, y: 3.0, w: 8.3, h: 2.4, fontSize: 9, border: { type: "solid", color: "D9DEDA", pt: 0.6 }, fill: "FFFFFF", margin: 0.05 });
  slide.addShape("roundRect", { x: 0.65, y: 6.25, w: 12, h: 0.55, fill: { color: "EEF6DF" }, line: { color: "86BC25" } });
  slide.addText("MANAGEMENT TAKEAWAY", { x: 0.82, y: 6.38, w: 1.05, h: 0.16, fontSize: 9, bold: true, margin: 0 });
  slide.addText("Example management takeaway", { x: 2.05, y: 6.34, w: 9.9, h: 0.24, fontSize: 12, bold: true, margin: 0 });
  const result = await pptx.write({ outputType: "uint8array", compression: true });
  return result instanceof Uint8Array ? result : new Uint8Array(result);
}

async function shapeTableTemplate() {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.addText("Example PMI Executive Status", { x: 0.55, y: 0.28, w: 11.8, h: 0.38, fontSize: 23, bold: true, margin: 0 });
  slide.addText("Example management message", { x: 0.55, y: 0.78, w: 11.5, h: 0.3, fontSize: 12.5, margin: 0 });
  [0.55, 4.05, 7.55].forEach((x, index) => {
    slide.addShape("roundRect", { x, y: 1.28, w: 3.1, h: 0.72, fill: { color: "F5F7F5" }, line: { color: "D9DEDA" } });
    slide.addText("EXAMPLE STATUS", { x: x + 0.18, y: 1.43, w: 1.35, h: 0.16, fontSize: 9, bold: true, margin: 0 });
    slide.addText(`Example ${index + 1}`, { x: x + 1.55, y: 1.39, w: 1.25, h: 0.24, fontSize: 17, bold: true, margin: 0 });
  });
  const columns = [0.55, 2.45, 3.35, 4.25];
  const widths = [1.9, 0.9, 0.9, 3.9];
  const headers = ["Workstream", "Status", "Progress", "Key update"];
  for (let row = 0; row < 4; row++) {
    for (let column = 0; column < 4; column++) {
      const y = 2.25 + row * 0.54;
      slide.addShape("rect", { x: columns[column], y, w: widths[column], h: 0.54, fill: { color: row ? "FFFFFF" : "202020" }, line: { color: "D9DEDA", pt: 0.5 } });
      if (row && column === 1) slide.addShape("ellipse", { x: columns[column] + 0.12, y: y + 0.2, w: 0.13, h: 0.13, fill: { color: "2E9B57" }, line: { transparency: 100 } });
      slide.addText(row ? `Example ${row}-${column}` : headers[column], { x: columns[column] + (row && column === 1 ? 0.3 : 0.08), y: y + 0.16, w: widths[column] - 0.16, h: 0.18, fontSize: 9, bold: !row, color: row ? "202020" : "FFFFFF", margin: 0 });
    }
  }
  slide.addShape("roundRect", { x: 8.45, y: 2.25, w: 3.7, h: 2.16, fill: { color: "F5F7F5" }, line: { color: "D9DEDA" } });
  slide.addText("Key management insights", { x: 8.7, y: 2.48, w: 3.15, h: 0.24, fontSize: 15, bold: true, margin: 0 });
  slide.addText("Example insight", { x: 8.7, y: 2.95, w: 3.1, h: 0.18, fontSize: 9, bold: true, margin: 0 });
  slide.addText("Example detail", { x: 8.7, y: 3.18, w: 3.1, h: 0.28, fontSize: 9, margin: 0 });
  slide.addText("Example insight 2", { x: 8.7, y: 3.56, w: 3.1, h: 0.18, fontSize: 9, bold: true, margin: 0 });
  slide.addText("Example detail 2", { x: 8.7, y: 3.79, w: 3.1, h: 0.28, fontSize: 9, margin: 0 });
  slide.addShape("roundRect", { x: 0.55, y: 5.75, w: 11.8, h: 0.72, fill: { color: "EEF6DF" }, line: { color: "86BC25" } });
  slide.addText("MANAGEMENT TAKEAWAY", { x: 0.78, y: 5.98, w: 1.05, h: 0.16, fontSize: 9, bold: true, margin: 0 });
  slide.addText("Example takeaway", { x: 2.1, y: 5.94, w: 9.6, h: 0.24, fontSize: 12, bold: true, margin: 0 });
  const result = await pptx.write({ outputType: "uint8array", compression: true });
  return result instanceof Uint8Array ? result : new Uint8Array(result);
}

test("resolves an explicit @filename without treating ordinary attachments as templates", () => {
  const evidence = { id: "evidence-1", fileName: "PMI_Data.xlsx", fileType: "xlsx", status: "extracted", excerpt: "Deal evidence" };
  const selected = resolveTemplateReference(
    "Using @Nike_adidas_PMI_Executive_Status.pptx create the same report for our PMI deal.",
    [templateSource, evidence],
  );
  assert.equal(selected?.sourceId, templateSource.id);
  assert.equal(templateOutputFormat(selected), "pptx");
  assert.equal(resolveTemplateReference("Create a PMI report from the attached files.", [templateSource, evidence]), null);
  assert.equal(messageMentionsTemplate("Use @File2.pptx", "File.pptx"), false);
  assert.equal(hasUnresolvedTemplateDirective("Use @Missing.pptx as the template"), true);
  assert.equal(hasUnresolvedTemplateDirective("Email finance@example.com"), false);
});

test("rejects ambiguous template mentions", () => {
  assert.throws(() => resolveTemplateReference("Use @first.pptx and @second.docx", [
    { ...templateSource, id: "first", fileName: "first.pptx" },
    { ...templateSource, id: "second", fileName: "second.docx", fileType: "docx" },
  ]), /Choose one template/);
});

test("keeps template facts outside evidence and lets the report determine slide count", () => {
  const prompt = buildPresentationPlanningPrompt({
    request: "Create the same report for our deal",
    audience: "Steering Committee",
    sources: [{ id: "deal", fileName: "Deal.xlsx", fileType: "xlsx", status: "extracted", excerpt: "Integration progress: 62%" }],
    history: [],
    template: templateSource,
  });
  assert.match(prompt, /source slide may be reused/i);
  assert.match(prompt, /Do not force the count to match the source deck/i);
  assert.match(prompt, /A template is not evidence/);
  assert.match(prompt, /Nike_adidas_PMI_Executive_Status\.pptx/);
  assert.match(prompt, /Integration progress: 62%/);
});

test("duplicates one template slide when the PMI storyline needs multiple slides", async () => {
  const sourceModel = {
    title: "Example status",
    audience: "Example Board",
    executiveSummary: "Example facts.",
    slides: [{ title: "Example status", layout: "cover", keyMessage: "Example facts.", items: [], sourceNotes: [] }],
  };
  const templateBytes = await renderPresentation(sourceModel);
  const targetModel = {
    title: "Orion risks and decisions",
    audience: "Steering Committee",
    executiveSummary: "Red risks require decisions.",
    slides: [
      { title: "Two red risks threaten cutover", layout: "risks", items: [{ label: "Cutover", detail: "Testing is blocked.", status: "red", sourceRefs: ["risk"] }], sourceNotes: [] },
      { title: "SteerCo decisions are required", layout: "decisions", items: [{ label: "Approve recovery plan", detail: "Decision required.", status: "amber", sourceRefs: ["risk"] }], sourceNotes: [] },
    ],
  };
  const output = fillPresentationTemplate(templateBytes, targetModel);
  const files = unzipSync(output);
  const slidePaths = Object.keys(files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
  assert.equal(slidePaths.length, 2);
  assert.match(strFromU8(files["ppt/slides/slide1.xml"]), /Two red risks threaten cutover/);
  assert.match(strFromU8(files["ppt/slides/slide2.xml"]), /SteerCo decisions are required/);
  assert.ok(files["ppt/slides/_rels/slide2.xml.rels"]);
  assert.ok(files["ppt/notesSlides/notesSlide2.xml"]);
  assert.match(strFromU8(files["ppt/slides/_rels/slide2.xml.rels"]), /notesSlides\/notesSlide2\.xml/);
  assert.match(strFromU8(files["ppt/notesSlides/_rels/notesSlide2.xml.rels"]), /slides\/slide2\.xml/);
  assert.equal((strFromU8(files["ppt/presentation.xml"]).match(/<p:sldId\b/g) ?? []).length, 2);
  assert.equal((strFromU8(files["ppt/_rels/presentation.xml.rels"]).match(/relationships\/slide"/g) ?? []).length, 2);
});

test("maps decisions into inherited cards and table rows without overlaying new text boxes", async () => {
  const templateBytes = await decisionTemplate();
  const layout = analyzePresentationTemplate(templateBytes);
  assert.equal(layout.slides.length, 1);
  assert.equal(layout.slides[0].regions.filter((region) => region.role === "card").length, 3);
  assert.equal(layout.slides[0].regions.find((region) => region.role === "table")?.maxItems, 5);
  assert.ok(layout.slides[0].maxItems >= 8);

  const target = {
    title: "SteerCo decisions should unblock IT, budget, people, and synergies",
    audience: "Steering Committee",
    executiveSummary: "Six decisions require action.",
    slides: [{
      title: "SteerCo decisions should unblock IT, budget, people, and synergies",
      templateSlide: 1,
      keyMessage: "Resolve the critical dependencies before the next reporting cycle.",
      layout: "decisions",
      items: Array.from({ length: 6 }, (_, index) => ({
        label: `Decision ${index + 1}`,
        detail: `Evidence-backed rationale ${index + 1}`,
        implication: `Management implication ${index + 1}`,
        owner: `Owner ${index + 1}`,
        deadline: `Week ${index + 1}`,
        status: index % 2 ? "amber" : "red",
        sourceRefs: [`source-${index + 1}`],
      })),
      sourceNotes: ["Risk register"],
    }],
  };
  const output = fillPresentationTemplate(templateBytes, target);
  const sourceFiles = unzipSync(templateBytes);
  const outputFiles = unzipSync(output);
  const sourceXml = strFromU8(sourceFiles["ppt/slides/slide1.xml"]);
  const outputXml = strFromU8(outputFiles["ppt/slides/slide1.xml"]);
  assert.equal((outputXml.match(/<p:sp>/g) ?? []).length, (sourceXml.match(/<p:sp>/g) ?? []).length, "renderer must not add overlay shapes");
  assert.equal((outputXml.match(/SteerCo decisions should unblock/g) ?? []).length, 1, "title is written exactly once");
  assert.match(outputXml, /Action required/);
  assert.match(outputXml, /Decision 1/);
  assert.match(outputXml, /Decision 6/);
  assert.match(outputXml, /Owner 4/);
  assert.doesNotMatch(outputXml, /Example decision|Example rationale|Example action|Example management takeaway/);
});

test("splits over-capacity content by duplicating and balancing the source layout", async () => {
  const templateBytes = await decisionTemplate();
  const target = {
    title: "Nine SteerCo decisions",
    audience: "Steering Committee",
    executiveSummary: "Nine decisions require action.",
    slides: [{
      title: "Nine SteerCo decisions",
      templateSlide: 1,
      keyMessage: "Resolve the dependencies.",
      layout: "decisions",
      items: Array.from({ length: 9 }, (_, index) => ({ label: `Decision ${index + 1}`, detail: `Rationale ${index + 1}`, status: "amber" })),
      sourceNotes: [],
    }],
  };
  const output = unzipSync(fillPresentationTemplate(templateBytes, target));
  const first = strFromU8(output["ppt/slides/slide1.xml"]);
  const second = strFromU8(output["ppt/slides/slide2.xml"]);
  assert.match(first, /Decision 5/);
  assert.doesNotMatch(first, /Decision 6/);
  assert.match(second, /Decision 6/);
  assert.match(second, /Decision 9/);
  assert.equal((first.match(/<p:sp>/g) ?? []).length, (second.match(/<p:sp>/g) ?? []).length);
});

test("maps pseudo-tables and ignores over-capacity fixed template labels", async () => {
  const templateBytes = await shapeTableTemplate();
  const layout = analyzePresentationTemplate(templateBytes).slides[0];
  assert.equal(layout.regions.find((region) => region.kind === "shape-table")?.maxItems, 3);
  assert.equal(layout.regions.find((region) => region.role === "panel")?.maxItems, 2);
  assert.ok(layout.regions.some((region) => region.role === "takeaway"));
  const target = {
    title: "SteerCo decisions should unblock IT, budget, people, and synergies",
    audience: "Steering Committee",
    executiveSummary: "Seven items need attention.",
    slides: [{
      title: "SteerCo decisions should unblock IT, budget, people, and synergies",
      templateSlide: 1,
      keyMessage: "Resolve the critical dependencies.",
      layout: "decisions",
      items: Array.from({ length: 7 }, (_, index) => ({ label: `Decision ${index + 1}`, detail: `Evidence ${index + 1}`, value: `${50 + index}%`, status: index % 2 ? "amber" : "red" })),
      sourceNotes: [],
    }],
  };
  const outputFiles = unzipSync(fillPresentationTemplate(templateBytes, target));
  const xml = strFromU8(outputFiles["ppt/slides/slide1.xml"]);
  assert.match(xml, /MANAGEMENT TAKEAWAY/);
  assert.match(xml, /Decision 7/);
  assert.match(xml, /Evidence 4/);
  assert.match(xml, /D92D20/);
  assert.doesNotMatch(xml, /Example PMI|Example management|Example insight|Example detail|Example takeaway|Example \d-\d/);
});

test("fills inherited PPTX slides while preserving the template package", async () => {
  const sourceModel = {
    title: "Example Company Status",
    audience: "Example Board",
    executiveSummary: "Example facts should be replaced.",
    slides: [
      { title: "Example Company Status", layout: "cover", keyMessage: "Example facts should be replaced.", items: [], sourceNotes: [] },
      { title: "Example delivery conclusion", layout: "cards", items: [{ label: "Example metric", value: "99%", detail: "Example only" }], sourceNotes: [] },
    ],
  };
  const templateBytes = await renderPresentation(sourceModel);
  const targetModel = {
    title: "Orion PMI Executive Status",
    audience: "Steering Committee",
    executiveSummary: "One dependency requires intervention.",
    slides: [
      { title: "Orion PMI Executive Status", layout: "cover", keyMessage: "One dependency requires intervention.", items: [], sourceNotes: [] },
      { title: "Cutover now defines the critical path", layout: "cards", items: [{ label: "Cutover readiness", value: "At risk", detail: "Testing sequence is unresolved.", sourceRefs: ["deal"] }], sourceNotes: ["Deal tracker"] },
    ],
  };
  const output = fillPresentationTemplate(templateBytes, targetModel);
  const files = unzipSync(output);
  assert.ok(files["ppt/slideMasters/slideMaster1.xml"]);
  const slideText = [1, 2].map((index) => strFromU8(files[`ppt/slides/slide${index}.xml`])).join("\n");
  assert.match(slideText, /Orion PMI Executive Status/);
  assert.match(slideText, /Cutover now defines the critical path/);
  assert.doesNotMatch(slideText, /Example metric|Example only|99%/);
});

test("fills native Excel, Word, and PDF template packages", async () => {
  const source = {
    title: "Example Company Report",
    audience: "Example Board",
    executiveSummary: "Example company facts.",
    sections: [{ name: "Example", title: "Example performance", type: "status", items: [{ label: "Example metric", value: "99%", detail: "Example only", status: "green" }] }],
  };
  const target = {
    title: "Orion PMI Report",
    audience: "Steering Committee",
    executiveSummary: "Cutover needs intervention.",
    sections: [{ name: "Readiness", title: "Cutover defines the critical path", type: "status", items: [{ label: "Cutover readiness", value: "At risk", detail: "Testing sequence is unresolved.", status: "red" }] }],
  };

  const sourceExcel = await renderExcelWorkbook(source, []);
  const excel = await fillExcelTemplate(sourceExcel.bytes, target);
  const excelFiles = unzipSync(excel);
  const excelText = Object.entries(excelFiles).filter(([path]) => /(?:sharedStrings|worksheets\/sheet\d+)\.xml$/.test(path)).map(([, bytes]) => strFromU8(bytes)).join("\n");
  assert.match(excelText, /Orion PMI Report/);

  const sourceWord = await renderWordDocument(source, []);
  const word = fillWordTemplate(new Uint8Array(sourceWord.bytes), target);
  const wordFiles = unzipSync(word);
  assert.match(strFromU8(wordFiles["word/document.xml"]), /Orion PMI Report/);
  assert.ok(wordFiles["word/header1.xml"]);

  const sourcePdf = await renderPdfReport(source, []);
  const pdf = await fillPdfTemplate(sourcePdf.bytes, target);
  const pdfDocument = await PDFDocument.load(pdf.bytes);
  assert.ok(pdfDocument.getPageCount() >= 2);
  assert.equal(pdfDocument.getTitle(), "Orion PMI Report");
});

test("preflights semantic capacity and preserves Excel formulas", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Decision log");
  sheet.getCell("A1").value = "Example PMI report";
  sheet.getCell("A1").font = { size: 22, bold: true };
  sheet.addRow([]);
  sheet.addRow(["Decision", "Owner", "By when", "Status", "Calculated"]);
  sheet.addRow(["Example decision", "Example owner", "Example date", "Amber", { formula: "1+1", result: 2 }]);
  const source = new Uint8Array(await workbook.xlsx.writeBuffer());
  const layout = await analyzeReportTemplate(source, "xlsx");
  assert.equal(layout.format, "xlsx");
  assert.equal(layout.worksheets[0].formulaCells, 1);
  assert.equal(layout.worksheets[0].tables[0].maxItems, 1);

  const target = {
    title: "Orion decisions",
    audience: "Steering Committee",
    executiveSummary: "One decision is required.",
    sections: [{ name: "Decisions", title: "Decision required", type: "decisions", items: [{ label: "Approve TSA exit", owner: "CIO", deadline: "Friday", status: "red", detail: "Cutover is blocked." }] }],
  };
  const output = await fillExcelTemplate(source, target);
  const rendered = new ExcelJS.Workbook();
  await rendered.xlsx.load(output);
  assert.equal(rendered.getWorksheet("Decision log").getCell("A1").value, "Orion decisions");
  assert.equal(rendered.getWorksheet("Decision log").getCell("A4").value, "Approve TSA exit");
  assert.equal(rendered.getWorksheet("Decision log").getCell("E4").value.formula, "1+1");
});

test("fills HTML placeholders or replaces one flow region without overlays", async () => {
  const target = {
    title: "Orion PMI Report",
    audience: "Steering Committee",
    executiveSummary: "Cutover needs intervention.",
    sections: [{ name: "Risks", title: "Red risks", type: "risks", items: [{ label: "Testing", detail: "Blocked", status: "red" }] }],
  };
  const placeholderTemplate = new TextEncoder().encode("<!doctype html><html><body><h1>{{ title }}</h1><p>{{ executive_summary }}</p></body></html>");
  const layout = await analyzeReportTemplate(placeholderTemplate, "html");
  assert.deepEqual(layout.placeholders, ["title", "executive_summary"]);
  const filled = fillHtmlTemplate(placeholderTemplate, target);
  assert.match(filled, /Orion PMI Report/);
  assert.doesNotMatch(filled, /{{/);

  const flowTemplate = new TextEncoder().encode("<!doctype html><html><head><style>.brand{color:green}</style></head><body><nav class=\"brand\">Brand</nav><main><h1>Example</h1></main><footer>Legal</footer></body></html>");
  const flow = fillHtmlTemplate(flowTemplate, target);
  assert.equal((flow.match(/<main\b/g) ?? []).length, 1);
  assert.match(flow, /<nav class="brand">Brand<\/nav>/);
  assert.match(flow, /<footer>Legal<\/footer>/);
  assert.doesNotMatch(flow, /<h1>Example<\/h1>/);
});

test("uses CSV headers as a tabular workbook template", async () => {
  const target = {
    title: "Orion decision log",
    audience: "Steering Committee",
    executiveSummary: "Decision required.",
    sections: [{ name: "Decisions", title: "Decision required", type: "decisions", items: [{ label: "Approve cutover", owner: "CIO", status: "red", detail: "Testing is blocked." }] }],
  };
  const output = await fillCsvTemplate(new TextEncoder().encode("Decision,Owner,Status,Detail\nExample,Example,Amber,Example"), target);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(output);
  const sheet = workbook.getWorksheet("PMI Report");
  assert.deepEqual(sheet.getRow(1).values.slice(1), ["Decision", "Owner", "Status", "Detail"]);
  assert.deepEqual(sheet.getRow(2).values.slice(1), ["Approve cutover", "CIO", "RED", "Testing is blocked."]);
});
