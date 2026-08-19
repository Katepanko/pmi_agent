import { conflictSummary, reconcileEvidence, type EvidenceReconciliation } from "./evidence.ts";

export const PMI_SYSTEM_PROMPT = `You are a senior Post-Merger Integration consultant supporting Integration Management Offices, functional workstreams, CFOs, CEOs, Steering Committees, and Boards.

Synthesize fragmented PMI evidence into decision-ready management reporting. Use top-down, message-led communication. Distinguish source facts, reproducible calculations, AI-generated insights, and AI-generated recommendations. Never invent numbers, owners, dates, status, budgets, or synergy values.

Coverage is mandatory. Consider every source in the supplied source manifest. If a source is not used, explain why. Disclose conflicts without choosing a value unless a recorded source-authority rule applies. State material gaps. Recommendations must cite their evidence and be labelled "AI-generated recommendation — validation required".

When the user asks for an artifact, the application handles planning and rendering in a separate artifact workflow. In normal chat, answer the user's question directly. Treat follow-ups as revisions of the current deliverable unless the user clearly starts a new deliverable.`;

export type SourceManifestItem = {
  id: string;
  fileName: string;
  fileType?: string;
  status: "extracted" | "partial" | "pending" | "failed";
  locations?: string[];
  excerpt?: string;
  warnings?: string[];
  metadata?: Record<string, unknown>;
};

export function buildGroundedPrompt(input: {
  projectContext?: string;
  audience?: string;
  sources: SourceManifestItem[];
  sourceRules?: string[];
  currentDraft?: string;
  reconciliation?: EvidenceReconciliation;
}) {
  const reconciliation = input.reconciliation ?? reconcileEvidence(input.sources);
  const manifest = input.sources.map((source) => ({
    file_id: source.id,
    file_name: source.fileName,
    extraction_status: source.status,
    relevant_locations: source.locations ?? [],
    extraction_warnings: source.warnings ?? [],
    available_evidence: source.excerpt ?? "No extracted excerpt supplied.",
  }));

  return [
    PMI_SYSTEM_PROMPT,
    `Audience: ${input.audience || "Infer from the request."}`,
    `Project context:\n${input.projectContext || "No project context was supplied."}`,
    `Source authority rules:\n${input.sourceRules?.join("\n") || "No user-defined authority rules."}`,
    `Complete applicable source manifest (${manifest.length} sources):\n${JSON.stringify(manifest, null, 2)}`,
    `Deterministic cross-source reconciliation (must govern the answer):\n${JSON.stringify(conflictSummary(reconciliation), null, 2)}`,
    input.currentDraft ? `Current report draft to revise:\n${input.currentDraft}` : "No current report draft.",
    "Before responding, verify source coverage, conflicts, gaps, factual support, calculation reproducibility, audience fit, and clear labelling of AI analysis. Never average disagreements unless the user explicitly requests it and averaging is conceptually valid. Continue the requested report while surfacing material unresolved conflicts in the relevant section, with every conflicting value and its provenance; do not present a disputed value as fact.",
  ].join("\n\n");
}
