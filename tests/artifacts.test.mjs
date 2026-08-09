import assert from "node:assert/strict";
import test from "node:test";
import { unzipSync } from "fflate";
import { PDFDocument } from "pdf-lib";
import { detectArtifactRequest } from "../app/lib/artifact-intent.ts";
import { renderExcelWorkbook } from "../app/lib/renderers/excel.ts";
import { renderWordDocument } from "../app/lib/renderers/word.ts";
import { renderPdfReport } from "../app/lib/renderers/pdf.ts";
import { renderHtmlDashboard } from "../app/lib/renderers/html.ts";

const report = {
  title: "Orion Integration Management Report",
  subtitle: "Decision-ready view",
  audience: "Steering Committee",
  reportingPeriod: "Week 6",
  executiveSummary: "Delivery remains broadly on track → two critical dependencies require management action this week.",
  sections: [
    {
      name: "Trajectory",
      title: "Delivery remains manageable if cross-functional dependencies are resolved",
      keyMessage: "The critical path is concentrated at the IT–Finance boundary.",
      type: "status",
      items: [
        { label: "Integration trajectory", value: "On track", detail: "Most planned milestones remain active.", implication: "Management attention can remain targeted.", status: "green", evidenceType: "fact", sourceRefs: ["src-plan"] },
        { label: "Cutover dependency → decision", detail: "The accountable owner is not evidenced.", implication: "The recovery path cannot yet be validated.", recommendation: "Assign one accountable executive.", status: "red", evidenceType: "gap", sourceRefs: ["src-risk"] },
      ],
      sourceNotes: ["Owner field is incomplete in the current risk register."],
    },
    {
      name: "Decisions",
      title: "Two decisions are required before the next reporting cycle",
      type: "decisions",
      items: [{ label: "Confirm cutover owner", detail: "Decision required", owner: "Not evidenced", deadline: "This week", status: "amber", evidenceType: "recommendation", sourceRefs: ["src-risk"] }],
    },
  ],
};

const sources = [
  { id: "src-plan", fileName: "Integration_Masterplan.xlsx", status: "extracted", excerpt: "Milestone and workstream records." },
  { id: "src-risk", fileName: "Risk_Register.xlsx", status: "partial", excerpt: "Risk and mitigation records.", warnings: ["Owner field incomplete"] },
];

test("routes explicit artifact requests and keeps normal chat conversational", () => {
  assert.equal(detectArtifactRequest("Create an Excel-based HR headcount and retention report."), "xlsx");
  assert.equal(detectArtifactRequest("I need a CFO Excel dashboard from this report."), "xlsx");
  assert.equal(detectArtifactRequest("Generate a Word executive summary."), "docx");
  assert.equal(detectArtifactRequest("Build a management PDF report."), "pdf");
  assert.equal(detectArtifactRequest("Create a standalone HTML dashboard."), "html");
  assert.equal(detectArtifactRequest("Generate a concise 8-slide PowerPoint deck."), "pptx");
  assert.equal(detectArtifactRequest("Summarize the current integration status."), null);
  assert.equal(detectArtifactRequest("Change the title and highlight the critical risks.", "pdf"), "pdf");
});

test("renders valid consulting-quality Excel, Word, PDF, and HTML packages", async () => {
  const excel = await renderExcelWorkbook(report, sources);
  const excelZip = unzipSync(excel.bytes);
  assert.ok(excel.sheetCount >= 4);
  assert.ok(excelZip["xl/workbook.xml"]);
  assert.ok(excelZip["xl/worksheets/sheet1.xml"]);

  const word = await renderWordDocument(report, sources);
  const wordZip = unzipSync(new Uint8Array(word.bytes));
  assert.equal(word.sectionCount, 3);
  assert.ok(wordZip["word/document.xml"]);
  assert.ok(wordZip["word/header1.xml"]);
  assert.ok(wordZip["word/footer1.xml"]);

  const pdf = await renderPdfReport(report, sources);
  const pdfDocument = await PDFDocument.load(pdf.bytes);
  assert.equal(pdfDocument.getPageCount(), pdf.pageCount);
  assert.ok(pdf.pageCount >= 1);

  const html = renderHtmlDashboard(report, sources);
  assert.match(html, /^<!doctype html>/i);
  assert.match(html, /Executive message/i);
  assert.match(html, /Evidence register/i);
  assert.match(html, /@media print/);
});
