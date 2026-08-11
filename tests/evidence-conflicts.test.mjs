import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, unzipSync } from "fflate";
import { reconcileEvidence } from "../app/lib/evidence.ts";
import { enforceConflictVisibility } from "../app/lib/conflict-guard.ts";
import { renderExcelWorkbook } from "../app/lib/renderers/excel.ts";
import { renderWordDocument } from "../app/lib/renderers/word.ts";
import { renderPdfReport } from "../app/lib/renderers/pdf.ts";
import { renderHtmlDashboard } from "../app/lib/renderers/html.ts";

function source(id, excerpt, metadata = {}) {
  return { id, fileName: `${id}.csv`, status: "extracted", excerpt, metadata, locations: [`${id}.csv → row 2`] };
}

test("same-period 78% / 66% / 70% is retained as an unresolved conflict without a selected value", () => {
  const result = reconcileEvidence([
    source("file-a", "Overall integration progress: 78%"),
    source("file-b", "Overall integration progress: 66%"),
    source("file-c", "Overall integration progress: 70%"),
  ]);
  assert.equal(result.conflicts.length, 1);
  assert.equal(result.conflicts[0].resolutionStatus, "unresolved_conflict");
  assert.equal(result.conflicts[0].selectedValue, undefined);
  assert.deepEqual(result.conflicts[0].observations.map((item) => item.value.numeric), [78, 66, 70]);
});

test("sequential reporting dates use the latest value and retain history", () => {
  const result = reconcileEvidence([
    source("week-1", "Metric: Overall integration progress | Value: 66% | Reporting Date: 2026-09-01"),
    source("week-2", "Metric: Overall integration progress | Value: 70% | Reporting Date: 2026-09-08"),
    source("week-3", "Metric: Overall integration progress | Value: 78% | Reporting Date: 2026-09-15"),
  ]);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.facts[0].resolutionStatus, "superseded");
  assert.equal(result.facts[0].selectedValue.raw, "78%");
  assert.equal(result.facts[0].observations.length, 3);
});

test("overall and IT progress are separate facts, not a conflict", () => {
  const result = reconcileEvidence([
    source("program", "Overall progress: 78%"),
    source("it", "IT progress: 66%"),
  ]);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.facts.length, 2);
});

test("70% and 70.1% agree within the percentage-point tolerance", () => {
  const result = reconcileEvidence([source("a", "Overall progress: 70%"), source("b", "Overall progress: 70.1%")]);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.facts[0].resolutionStatus, "confirmed");
  assert.match(result.facts[0].resolutionReason, /tolerance/i);
});

test("an explicit user correction resolves a conflict while preserving observations", () => {
  const result = reconcileEvidence(
    [source("a", "Overall integration progress: 66%"), source("b", "Overall integration progress: 78%")],
    { userStatements: ["Overall integration progress: 78% is the correct approved figure."] },
  );
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.facts[0].resolutionStatus, "resolved_by_user");
  assert.equal(result.facts[0].selectedValue.raw, "78%");
  assert.equal(result.facts[0].observations.length, 2);
});

test("the newer approved master observation supersedes the older one with provenance", () => {
  const result = reconcileEvidence([
    source("old-master", "Metric: Overall progress | Value: 66% | Reporting Date: 2026-09-01", { designation: "master tracker", approved: true }),
    source("new-master", "Metric: Overall progress | Value: 70% | Reporting Date: 2026-09-08", { designation: "master tracker", approved: true }),
  ]);
  assert.equal(result.facts[0].resolutionStatus, "superseded");
  assert.equal(result.facts[0].selectedValue.raw, "70%");
  assert.equal(result.facts[0].selectedObservation.sourceId, "new-master");
});

test("conflicting synergy values are forced into a Finance/SteerCo report model", () => {
  const reconciliation = reconcileEvidence([source("finance-a", "Synergy realization: €20m"), source("finance-b", "Synergy realization: €14m")]);
  const guarded = enforceConflictVisibility({
    title: "SteerCo report", audience: "Steering Committee", executiveSummary: "Integration update", sections: [{ name: "Summary", title: "Current position", type: "summary", items: [] }],
  }, reconciliation);
  const section = guarded.sections.at(-1);
  assert.equal(section.name, "Finance conflicts");
  assert.match(JSON.stringify(section), /€20m/);
  assert.match(JSON.stringify(section), /€14m/);
  assert.match(JSON.stringify(section), /finance-a\.csv/);
});

test("one guarded conflict model remains visible in PPTX/PDF/Word/Excel/HTML rendering", async () => {
  const sources = [source("a", "Overall integration progress: 78%"), source("b", "Overall integration progress: 66%")];
  const reconciliation = reconcileEvidence(sources);
  const baseReport = { title: "PMI report", audience: "Management", executiveSummary: "Status update", sections: [{ name: "Summary", title: "Status", type: "summary", items: [] }] };
  const report = enforceConflictVisibility(baseReport, reconciliation);
  const presentation = enforceConflictVisibility({ title: "PMI deck", audience: "SteerCo", executiveSummary: "Status update", slides: [{ title: "PMI deck", layout: "cover", items: [] }] }, reconciliation);

  const [pdf, docx, xlsx, htmlText] = await Promise.all([
    renderPdfReport(report, sources),
    renderWordDocument(report, sources),
    renderExcelWorkbook(report, sources),
    Promise.resolve(renderHtmlDashboard(report, sources)),
  ]);
  const wordText = strFromU8(unzipSync(docx.bytes)["word/document.xml"]);
  const excelText = Object.entries(unzipSync(xlsx.bytes)).filter(([name]) => /xl\/(sharedStrings|worksheets)/.test(name)).map(([, bytes]) => strFromU8(bytes)).join(" ");
  assert.match(JSON.stringify(presentation), /Data conflict/); // PPTX consumes this guarded presentation model.
  assert.match(wordText, /Data conflict/);
  assert.match(excelText, /Data conflict/);
  assert.match(htmlText, /Data conflict/);
  assert.ok(pdf.bytes.byteLength > 1_000);
  assert.match(JSON.stringify(report), /Data conflict/); // PDF consumes this same guarded model.
});
