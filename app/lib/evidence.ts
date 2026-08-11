import type { SourceManifestItem } from "./pmi-prompt";

export type EvidenceValue = {
  raw: string;
  kind: "number" | "percentage" | "date" | "status" | "text";
  numeric?: number;
  unit?: string;
};

export type SourceObservation = {
  metricKey: string;
  metricName: string;
  scope?: string;
  reportingDate?: string;
  reportingPeriod?: string;
  entity?: string;
  value: EvidenceValue;
  sourceId: string;
  sourceName: string;
  location?: string;
  authority: "explicit_master" | "approved" | "ordinary";
  assertionText: string;
};

export type ConflictMateriality = "critical" | "relevant" | "minor";
export type ResolutionStatus = "confirmed" | "superseded" | "resolved_by_user" | "unresolved_conflict";

export type ReconciledFact = {
  metricKey: string;
  metricName: string;
  observations: SourceObservation[];
  resolutionStatus: ResolutionStatus;
  selectedValue?: EvidenceValue;
  selectedObservation?: SourceObservation;
  resolutionReason?: string;
  materiality: ConflictMateriality;
};

export type EvidenceReconciliation = {
  facts: ReconciledFact[];
  conflicts: ReconciledFact[];
  observations: SourceObservation[];
};

const VALUE = String.raw`(?:[-+]?[$€£]?\s*\d[\d,.]*\s*(?:%|percent|percentage points?|pp|[kmb]|million|billion)?|green|amber|red|on track|at risk|critical|complete|completed|not started|blocked|delayed|[A-Z][A-Za-z -]{1,30})`;
const PAIR = new RegExp(String.raw`^\s*(?:[-*•]\s*)?([^:=|]{2,100}?)\s*(?:=|:)\s*(${VALUE})(?:\s*(?:\||;|—|-).*)?$`, "i");
const RECORD_FIELD = /(?:^|\|)\s*([^:|]{1,50}):\s*([^|]*?)\s*(?=\||$)/g;
const DATE_PATTERN = /\b(20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]20\d{2}|\d{1,2}\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+20\d{2})\b/i;

function normalizedWords(value: string) {
  return value.toLowerCase().replace(/[_-]+/g, " ").replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");
}

function metricKey(name: string) {
  return normalizedWords(name)
    .replace(/\b(kpi|metric|value|figure|current|actual)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDate(value?: string) {
  if (!value) return undefined;
  const normalized = value.replace(/\//g, "-");
  const time = Date.parse(normalized);
  return Number.isFinite(time) ? new Date(time).toISOString().slice(0, 10) : undefined;
}

function parseValue(rawValue: string): EvidenceValue {
  const raw = rawValue.trim().replace(/[.;,]$/, "");
  const lower = raw.toLowerCase();
  const numberMatch = raw.match(/[-+]?[$€£]?\s*([\d,.]+)/);
  if (numberMatch) {
    const numeric = Number(numberMatch[1].replace(/,/g, ""));
    if (Number.isFinite(numeric)) {
      if (/%|percent/.test(lower)) return { raw, kind: "percentage", numeric, unit: "%" };
      const currency = raw.match(/[$€£]/)?.[0];
      const scale = lower.match(/\b(k|m|b|million|billion)\b/)?.[0];
      return { raw, kind: "number", numeric, unit: [currency, scale].filter(Boolean).join("") || undefined };
    }
  }
  if (/^(green|amber|red|on track|at risk|critical|complete|completed|not started|blocked|delayed)$/i.test(raw)) return { raw, kind: "status" };
  if (parseDate(raw)) return { raw, kind: "date" };
  return { raw, kind: "text" };
}

function authorityFor(source: SourceManifestItem, text: string): SourceObservation["authority"] {
  const metadata = source.metadata ?? {};
  if (metadata.isAuthoritative === true || /\b(authoritative|source of truth|master tracker)\b/i.test(String(metadata.designation ?? ""))) return "explicit_master";
  if (metadata.approved === true || /\bapproved\b/i.test(String(metadata.approvalStatus ?? ""))) return "approved";
  if (/\b(approved (?:master |steerco)|authoritative (?:master |source)|designated source of truth)\b/i.test(text)) return /master|source of truth/i.test(text) ? "explicit_master" : "approved";
  return "ordinary";
}

function observation(source: SourceManifestItem, name: string, rawValue: string, assertionText: string, fields: Record<string, string> = {}, lineNumber?: number): SourceObservation | null {
  const key = metricKey(name);
  if (!key || /^(source|file|sheet|row|page|slide|owner|reporting date|date|period|scope|entity|unit)$/.test(key)) return null;
  const dateText = fields["reporting date"] || fields.date || assertionText.match(DATE_PATTERN)?.[0] || String(source.metadata?.reportingDate ?? "");
  return {
    metricKey: key,
    metricName: name.trim(),
    scope: fields.scope?.trim() || undefined,
    reportingDate: parseDate(dateText),
    reportingPeriod: fields.period?.trim() || fields["reporting period"]?.trim() || undefined,
    entity: fields.entity?.trim() || undefined,
    value: parseValue(rawValue),
    sourceId: source.id,
    sourceName: source.fileName,
    location: source.locations?.[lineNumber ? lineNumber - 1 : 0] || (lineNumber ? `${source.fileName} → extracted line ${lineNumber}` : source.fileName),
    authority: authorityFor(source, `${assertionText} ${source.excerpt ?? ""}`),
    assertionText,
  };
}

export function extractSourceObservations(sources: SourceManifestItem[]): SourceObservation[] {
  const observations: SourceObservation[] = [];
  for (const source of sources) {
    const lines = (source.excerpt ?? "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    lines.forEach((line, index) => {
      const fields: Record<string, string> = {};
      for (const match of line.matchAll(RECORD_FIELD)) fields[normalizedWords(match[1])] = match[2].trim();
      const name = fields.metric || fields.kpi || fields.measure || fields.indicator;
      const rawValue = fields.value || fields.actual || fields.status || fields.progress || fields.amount;
      if (name && rawValue) {
        const found = observation(source, name, rawValue, line, fields, index + 1);
        if (found) observations.push(found);
        return;
      }
      const match = line.match(PAIR);
      if (match) {
        const found = observation(source, match[1], match[2], line, fields, index + 1);
        if (found) observations.push(found);
      }
    });
  }
  return observations;
}

function sameValue(a: EvidenceValue, b: EvidenceValue, percentageTolerance: number) {
  if (a.kind !== b.kind) return normalizedWords(a.raw) === normalizedWords(b.raw);
  if (a.numeric !== undefined && b.numeric !== undefined) {
    if ((a.unit ?? "") !== (b.unit ?? "")) return false;
    const tolerance = a.kind === "percentage" ? percentageTolerance : Math.max(0.000001, Math.max(Math.abs(a.numeric), Math.abs(b.numeric)) * 0.001);
    return Math.abs(a.numeric - b.numeric) <= tolerance;
  }
  return normalizedWords(a.raw) === normalizedWords(b.raw);
}

function comparableKey(item: SourceObservation) {
  return [item.metricKey, normalizedWords(item.scope ?? ""), normalizedWords(item.reportingPeriod ?? ""), normalizedWords(item.entity ?? ""), item.value.unit ?? ""].join("|");
}

function materiality(name: string, equivalent: boolean): ConflictMateriality {
  if (equivalent) return "minor";
  if (/\b(overall|program|synerg|budget|actual|financial|milestone|day 1|day-1|readiness|major risk)\b/i.test(name)) return "critical";
  if (/\b(progress|owner|status|date|workstream|risk)\b/i.test(name)) return "relevant";
  return "relevant";
}

function userResolution(metricName: string, history: string[], allowContextualCorrection: boolean) {
  const candidates = history.filter((entry) => /\b(correct|approved|use|authoritative|source of truth)\b/i.test(entry));
  for (const entry of candidates.reverse()) {
    if (!normalizedWords(entry).includes(normalizedWords(metricName)) && !(allowContextualCorrection && /\b(correct approved figure|correct figure)\b/i.test(entry))) continue;
    const value = entry.match(/[-+]?[$€£]?\s*\d[\d,.]*\s*(?:%|percent|[kmb]|million|billion)?/i)?.[0];
    if (value) return parseValue(value);
  }
  return undefined;
}

export function reconcileEvidence(sources: SourceManifestItem[], options: { percentageTolerance?: number; userStatements?: string[]; authorityRules?: string[] } = {}): EvidenceReconciliation {
  const ruledSources = sources.map((source) => {
    const matchingRule = (options.authorityRules ?? []).find((rule) => {
      const normalizedRule = normalizedWords(rule);
      return (normalizedRule.includes(normalizedWords(source.fileName)) || normalizedRule.includes(normalizedWords(source.id))) && /\b(authoritative|master|source of truth|approved|priority)\b/i.test(rule);
    });
    return matchingRule ? { ...source, metadata: { ...source.metadata, isAuthoritative: true, designation: matchingRule } } : source;
  });
  const observations = extractSourceObservations(ruledSources);
  const groups = new Map<string, SourceObservation[]>();
  observations.forEach((item) => groups.set(comparableKey(item), [...(groups.get(comparableKey(item)) ?? []), item]));
  const facts: ReconciledFact[] = [];
  for (const group of groups.values()) {
    const tolerance = options.percentageTolerance ?? 0.2;
    const equivalent = group.every((item) => sameValue(group[0].value, item.value, tolerance));
    const base = { metricKey: group[0].metricKey, metricName: group[0].metricName, observations: group, materiality: materiality(group[0].metricName, equivalent) };
    if (equivalent) {
      facts.push({ ...base, resolutionStatus: "confirmed", selectedValue: group[group.length - 1].value, selectedObservation: group[group.length - 1], resolutionReason: group.length > 1 ? `Values agree within tolerance (${tolerance} percentage points for percentages).` : "Single source assertion." });
      continue;
    }
    const correction = userResolution(group[0].metricName, options.userStatements ?? [], groups.size === 1);
    if (correction) {
      const selected = group.find((item) => sameValue(item.value, correction, 0.000001));
      facts.push({ ...base, resolutionStatus: "resolved_by_user", selectedValue: correction, selectedObservation: selected, resolutionReason: "The user explicitly identified the approved value; the conflicting source history is retained." });
      continue;
    }
    const dated = group.filter((item) => item.reportingDate).sort((a, b) => a.reportingDate!.localeCompare(b.reportingDate!));
    if (dated.length === group.length && new Set(dated.map((item) => item.reportingDate)).size === group.length) {
      const latest = dated[dated.length - 1];
      facts.push({ ...base, resolutionStatus: "superseded", selectedValue: latest.value, selectedObservation: latest, resolutionReason: "The latest dated observation supersedes earlier reporting periods; history is retained." });
      continue;
    }
    const highestAuthority = Math.max(...group.map((item) => item.authority === "explicit_master" ? 2 : item.authority === "approved" ? 1 : 0));
    const authoritative = group.filter((item) => (item.authority === "explicit_master" ? 2 : item.authority === "approved" ? 1 : 0) === highestAuthority);
    if (highestAuthority > 0 && authoritative.length === 1) {
      const selected = authoritative[0];
      facts.push({ ...base, resolutionStatus: "superseded", selectedValue: selected.value, selectedObservation: selected, resolutionReason: "Selected from the sole explicitly authoritative or approved source; no file-type hierarchy was inferred." });
      continue;
    }
    facts.push({ ...base, resolutionStatus: "unresolved_conflict", resolutionReason: "Values differ materially and date, scope, authority, and user-correction evidence do not resolve the discrepancy." });
  }
  return { facts, conflicts: facts.filter((fact) => fact.resolutionStatus === "unresolved_conflict"), observations };
}

export function conflictSummary(reconciliation: EvidenceReconciliation) {
  if (!reconciliation.conflicts.length) return "No unresolved cross-source conflicts were deterministically detected in the supplied excerpts.";
  return reconciliation.conflicts.map((fact) => ({
    metric: fact.metricName,
    materiality: fact.materiality,
    resolution_status: fact.resolutionStatus,
    selected_value: null,
    observations: fact.observations.map((item) => ({ value: item.value.raw, source_id: item.sourceId, source_file: item.sourceName, location: item.location, reporting_date: item.reportingDate })),
  }));
}
