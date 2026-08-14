import type { ArtifactFormat } from "./artifact-intent.ts";

const nullableString = { type: ["string", "null"] } as const;
const sourceRefsSchema = { type: "array", items: { type: "string" } } as const;

const presentationItemSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    label: { type: "string" },
    value: nullableString,
    detail: nullableString,
    implication: nullableString,
    owner: nullableString,
    deadline: nullableString,
    status: { type: "string", enum: ["green", "amber", "red", "neutral"] },
    evidenceType: { type: "string", enum: ["fact", "calculation", "inference", "recommendation", "gap", "conflict"] },
    sourceRefs: sourceRefsSchema,
  },
  required: ["label", "value", "detail", "implication", "owner", "deadline", "status", "evidenceType", "sourceRefs"],
} as const;

const presentationSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    title: { type: "string" },
    subtitle: nullableString,
    projectName: nullableString,
    location: nullableString,
    date: nullableString,
    audience: { type: "string" },
    executiveSummary: { type: "string" },
    slides: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          title: { type: "string" },
          kicker: nullableString,
          keyMessage: nullableString,
          layout: { type: "string", enum: ["cover", "summary", "trajectory", "risks", "synergies", "decisions", "timeline", "comparison", "cards"] },
          items: { type: "array", items: presentationItemSchema },
          sourceNotes: { type: "array", items: { type: "string" } },
        },
        required: ["title", "kicker", "keyMessage", "layout", "items", "sourceNotes"],
      },
    },
  },
  required: ["title", "subtitle", "projectName", "location", "date", "audience", "executiveSummary", "slides"],
} as const;

const reportItemSchema = {
  ...presentationItemSchema,
  properties: { ...presentationItemSchema.properties, recommendation: nullableString },
  required: [...presentationItemSchema.required, "recommendation"],
} as const;

const consultingReportSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    title: { type: "string" },
    subtitle: nullableString,
    audience: { type: "string" },
    reportingPeriod: nullableString,
    executiveSummary: { type: "string" },
    sections: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          name: { type: "string" },
          title: { type: "string" },
          keyMessage: nullableString,
          type: { type: "string", enum: ["summary", "kpi", "status", "milestones", "risks", "synergies", "actions", "decisions", "sources", "detail"] },
          items: { type: "array", items: reportItemSchema },
          sourceNotes: { type: "array", items: { type: "string" } },
        },
        required: ["name", "title", "keyMessage", "type", "items", "sourceNotes"],
      },
    },
  },
  required: ["title", "subtitle", "audience", "reportingPeriod", "executiveSummary", "sections"],
} as const;

export function artifactStructuredOutput(format: ArtifactFormat) {
  return {
    name: format === "pptx" ? "pmi_presentation" : "pmi_consulting_report",
    schema: (format === "pptx" ? presentationSchema : consultingReportSchema) as unknown as Record<string, unknown>,
  };
}
