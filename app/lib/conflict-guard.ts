import type { ArtifactContentModel } from "./artifact";
import type { EvidenceReconciliation } from "./evidence";

export function enforceConflictVisibility(model: ArtifactContentModel, reconciliation: EvidenceReconciliation): ArtifactContentModel {
  if (!reconciliation.conflicts.length) return model;
  const conflictItems = reconciliation.conflicts.map((conflict) => ({
    label: `Data conflict — ${conflict.metricName}`,
    value: conflict.observations.map((item) => item.value.raw).join(" / "),
    detail: conflict.observations.map((item) => `${item.value.raw} — ${item.sourceName}${item.location ? ` (${item.location})` : ""}`).join("; "),
    implication: conflict.materiality === "critical" ? "Reconcile before using this figure for management decisions." : "Confirm the authoritative value before relying on this field.",
    status: conflict.materiality === "critical" ? "red" as const : "amber" as const,
    evidenceType: "conflict" as const,
    sourceRefs: [...new Set(conflict.observations.map((item) => item.sourceId))],
  }));
  if ("slides" in model) {
    const existingText = JSON.stringify(model).toLowerCase();
    const missing = conflictItems.filter((item) => !existingText.includes(item.label.toLowerCase()));
    if (!missing.length) return model;
    return { ...model, slides: [...model.slides, { title: "Source conflicts require reconciliation before decisions", kicker: "Data quality", keyMessage: "Disputed values remain unresolved and have not been consolidated.", layout: "cards", items: missing.slice(0, 6), sourceNotes: ["No file-type authority hierarchy was applied; source assertions remain traceable."] }] };
  }
  const existingText = JSON.stringify(model).toLowerCase();
  const missing = conflictItems.filter((item) => !existingText.includes(item.label.toLowerCase()));
  if (!missing.length) return model;
  const isFinance = missing.some((item) => /synerg|budget|financial|actual/i.test(item.label));
  return { ...model, sections: [...model.sections, { name: isFinance ? "Finance conflicts" : "Data quality", title: "Source conflicts must be reconciled before management reliance", keyMessage: "Disputed values are shown without averaging or silent selection.", type: isFinance ? "synergies" : "detail", items: missing, sourceNotes: ["No file-type authority hierarchy was applied; all assertions and source locations are retained."] }] };
}
