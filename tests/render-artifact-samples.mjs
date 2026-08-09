import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { renderExcelWorkbook } from "../app/lib/renderers/excel.ts";
import { renderWordDocument } from "../app/lib/renderers/word.ts";
import { renderPdfReport } from "../app/lib/renderers/pdf.ts";
import { renderHtmlDashboard } from "../app/lib/renderers/html.ts";

const output = resolve(process.argv[2] ?? "/private/tmp/pmi-artifact-qa");
const model = {
  title: "Orion / Nova Integration — Executive Report",
  subtitle: "Steering Committee decision pack",
  audience: "Steering Committee",
  reportingPeriod: "Week 6 after Day 1",
  executiveSummary: "Integration delivery remains broadly on track → cutover accountability and synergy timing require executive resolution this week.",
  sections: [
    { name: "Trajectory", title: "Execution remains manageable, with pressure concentrated at the IT–Finance boundary", keyMessage: "Most milestones remain active; the only critical-path exposure is an unresolved cross-functional cutover dependency.", type: "status", items: [
      { label: "Overall trajectory", value: "On track", detail: "The integrated plan shows most workstreams continuing against the current baseline.", implication: "Management attention can remain focused on the small number of exceptions.", status: "green", evidenceType: "fact", sourceRefs: ["src-plan"] },
      { label: "IT–Finance cutover", value: "Owner gap", detail: "The risk register does not evidence one accountable executive for the dependency.", implication: "A credible recovery path cannot be locked until accountability is resolved.", recommendation: "Nominate one executive owner and publish the reconciled dependency plan.", status: "red", evidenceType: "gap", sourceRefs: ["src-risk"] },
      { label: "Data readiness", value: "At risk", detail: "Validation activity is active, but completion evidence is partial.", implication: "Late defect discovery could compress the cutover window.", status: "amber", evidenceType: "inference", sourceRefs: ["src-plan"] },
    ], sourceNotes: ["Cutover ownership is missing in the current risk-register extract."] },
    { name: "Value delivery", title: "Synergy initiatives remain active, but timing confidence is not yet decision-grade", keyMessage: "Value-delivery reporting should remain internal until tracker timing is reconciled.", type: "synergies", items: [
      { label: "Confirmed initiatives", value: "Active", detail: "The tracker contains active initiatives and accountable workstreams.", implication: "The value pipeline remains intact.", status: "green", evidenceType: "fact", sourceRefs: ["src-synergy"] },
      { label: "Realization timing", value: "Conflict", detail: "The weekly update and baseline tracker contain inconsistent timing references.", implication: "External reporting would imply more precision than the evidence supports.", recommendation: "Confirm the authoritative baseline before the next SteerCo.", status: "amber", evidenceType: "recommendation", sourceRefs: ["src-synergy", "src-weekly"] },
    ] },
    { name: "Decisions", title: "Two decisions this week will protect the integrated plan and value narrative", type: "decisions", items: [
      { label: "Confirm cutover accountability", detail: "Assign one accountable executive for the IT–Finance dependency.", owner: "SteerCo", deadline: "This week", status: "red", evidenceType: "recommendation", sourceRefs: ["src-risk"] },
      { label: "Approve source authority", detail: "Confirm whether the baseline synergy tracker supersedes the weekly narrative.", owner: "CFO", deadline: "Before next SteerCo", status: "amber", evidenceType: "recommendation", sourceRefs: ["src-synergy", "src-weekly"] },
    ] },
  ],
};
const sources = [
  { id: "src-plan", fileName: "Integration_Masterplan.xlsx", status: "extracted", excerpt: "Milestones, workstreams, and dependencies." },
  { id: "src-risk", fileName: "Risk_Register.xlsx", status: "partial", excerpt: "Risk impact and mitigation records.", warnings: ["Accountable owner is missing for one critical dependency."] },
  { id: "src-synergy", fileName: "Synergy_Tracker.xlsx", status: "extracted", excerpt: "Initiatives, targets, timing, and owners." },
  { id: "src-weekly", fileName: "Weekly_Update.docx", status: "extracted", excerpt: "Workstream updates and decisions." },
];

await mkdir(output, { recursive: true });
const [xlsx, docx, pdf] = await Promise.all([renderExcelWorkbook(model, sources), renderWordDocument(model, sources), renderPdfReport(model, sources)]);
await Promise.all([
  writeFile(resolve(output, "consulting-report.xlsx"), xlsx.bytes),
  writeFile(resolve(output, "consulting-report.docx"), docx.bytes),
  writeFile(resolve(output, "consulting-report.pdf"), pdf.bytes),
  writeFile(resolve(output, "consulting-dashboard.html"), renderHtmlDashboard(model, sources)),
]);
console.log(output);
