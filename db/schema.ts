import { index, integer, sqliteTable, text, uniqueIndex } from "drizzle-orm/sqlite-core";

const timestamps = {
  createdAt: text("created_at").notNull(),
  updatedAt: text("updated_at").notNull(),
};

export const projects = sqliteTable(
  "projects",
  {
    id: text("id").primaryKey(),
    userId: text("user_id").notNull(),
    name: text("name").notNull(),
    description: text("description").notNull().default(""),
    companies: text("companies").notNull().default(""),
    dealRationale: text("deal_rationale").notNull().default(""),
    integrationType: text("integration_type").notNull().default(""),
    dayOneDate: text("day_one_date"),
    objectives: text("objectives").notNull().default(""),
    synergyTargets: text("synergy_targets").notNull().default(""),
    governance: text("governance").notNull().default(""),
    terminology: text("terminology").notNull().default(""),
    reportingExpectations: text("reporting_expectations").notNull().default(""),
    instructions: text("instructions").notNull().default(""),
    icon: text("icon").notNull().default("layers"),
    archivedAt: text("archived_at"),
    ...timestamps,
  },
  (table) => [index("idx_projects_user_archived").on(table.userId, table.archivedAt)],
);

export const chats = sqliteTable(
  "chats",
  {
    id: text("id").primaryKey(),
    userId: text("user_id").notNull(),
    projectId: text("project_id").references(() => projects.id, { onDelete: "cascade" }),
    title: text("title").notNull(),
    modelKey: text("model_key").notNull().default("openai-gpt56"),
    audience: text("audience").notNull().default("Steering Committee"),
    archivedAt: text("archived_at"),
    ...timestamps,
  },
  (table) => [
    index("idx_chats_user_project").on(table.userId, table.projectId),
    index("idx_chats_user_updated").on(table.userId, table.updatedAt),
  ],
);

export const messages = sqliteTable(
  "messages",
  {
    id: text("id").primaryKey(),
    chatId: text("chat_id").notNull().references(() => chats.id, { onDelete: "cascade" }),
    role: text("role", { enum: ["user", "assistant", "system"] }).notNull(),
    content: text("content").notNull(),
    classification: text("classification").notNull().default("conversation"),
    modelKey: text("model_key"),
    sourceCoverageJson: text("source_coverage_json"),
    stoppedAt: text("stopped_at"),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("idx_messages_chat_created").on(table.chatId, table.createdAt)],
);

export const sources = sqliteTable(
  "sources",
  {
    id: text("id").primaryKey(),
    userId: text("user_id").notNull(),
    projectId: text("project_id").references(() => projects.id, { onDelete: "cascade" }),
    chatId: text("chat_id").notNull().references(() => chats.id, { onDelete: "cascade" }),
    messageId: text("message_id").references(() => messages.id, { onDelete: "set null" }),
    fileName: text("file_name").notNull(),
    fileType: text("file_type").notNull(),
    mediaType: text("media_type").notNull(),
    sizeBytes: integer("size_bytes").notNull(),
    objectKey: text("object_key").notNull(),
    checksum: text("checksum"),
    extractionStatus: text("extraction_status").notNull().default("pending"),
    extractionWarningsJson: text("extraction_warnings_json").notNull().default("[]"),
    metadataJson: text("metadata_json").notNull().default("{}"),
    createdAt: text("created_at").notNull(),
    updatedAt: text("updated_at").notNull(),
  },
  (table) => [
    index("idx_sources_chat_status").on(table.chatId, table.extractionStatus),
    index("idx_sources_project").on(table.projectId),
  ],
);

export const sourceSegments = sqliteTable(
  "source_segments",
  {
    id: text("id").primaryKey(),
    sourceId: text("source_id").notNull().references(() => sources.id, { onDelete: "cascade" }),
    ordinal: integer("ordinal").notNull(),
    location: text("location").notNull(),
    kind: text("kind").notNull(),
    content: text("content").notNull(),
    structuredJson: text("structured_json"),
    confidence: text("confidence").notNull().default("high"),
    createdAt: text("created_at").notNull(),
  },
  (table) => [
    uniqueIndex("idx_source_segments_source_ordinal").on(table.sourceId, table.ordinal),
    index("idx_source_segments_kind").on(table.kind),
  ],
);

export const knowledgeVersions = sqliteTable(
  "knowledge_versions",
  {
    id: text("id").primaryKey(),
    projectId: text("project_id").references(() => projects.id, { onDelete: "cascade" }),
    chatId: text("chat_id").references(() => chats.id, { onDelete: "cascade" }),
    version: integer("version").notNull(),
    scope: text("scope", { enum: ["project", "chat"] }).notNull(),
    contentJson: text("content_json").notNull(),
    createdAt: text("created_at").notNull(),
  },
  (table) => [
    uniqueIndex("idx_knowledge_scope_version").on(table.scope, table.projectId, table.chatId, table.version),
  ],
);

export const sourcePriorityRules = sqliteTable(
  "source_priority_rules",
  {
    id: text("id").primaryKey(),
    projectId: text("project_id").references(() => projects.id, { onDelete: "cascade" }),
    chatId: text("chat_id").references(() => chats.id, { onDelete: "cascade" }),
    scope: text("scope", { enum: ["project", "chat"] }).notNull(),
    sourceId: text("source_id").notNull().references(() => sources.id, { onDelete: "cascade" }),
    appliesTo: text("applies_to").notNull(),
    instruction: text("instruction").notNull(),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("idx_source_rules_chat").on(table.chatId)],
);

export const reportDrafts = sqliteTable(
  "report_drafts",
  {
    id: text("id").primaryKey(),
    projectId: text("project_id").references(() => projects.id, { onDelete: "cascade" }),
    chatId: text("chat_id").notNull().references(() => chats.id, { onDelete: "cascade" }),
    title: text("title").notNull(),
    audience: text("audience").notNull(),
    reportType: text("report_type").notNull(),
    requestedFormat: text("requested_format"),
    currentVersion: integer("current_version").notNull().default(1),
    basedOnKnowledgeVersion: integer("based_on_knowledge_version").notNull().default(1),
    createdAt: text("created_at").notNull(),
    updatedAt: text("updated_at").notNull(),
  },
  (table) => [index("idx_report_drafts_chat").on(table.chatId, table.updatedAt)],
);

export const reportDraftVersions = sqliteTable(
  "report_draft_versions",
  {
    id: text("id").primaryKey(),
    draftId: text("draft_id").notNull().references(() => reportDrafts.id, { onDelete: "cascade" }),
    version: integer("version").notNull(),
    content: text("content").notNull(),
    sectionsJson: text("sections_json").notNull().default("[]"),
    sourcesJson: text("sources_json").notNull().default("[]"),
    createdAt: text("created_at").notNull(),
  },
  (table) => [uniqueIndex("idx_draft_version").on(table.draftId, table.version)],
);

export const artifacts = sqliteTable(
  "artifacts",
  {
    id: text("id").primaryKey(),
    userId: text("user_id").notNull(),
    projectId: text("project_id").references(() => projects.id, { onDelete: "set null" }),
    chatId: text("chat_id").notNull().references(() => chats.id, { onDelete: "cascade" }),
    messageId: text("message_id").notNull().references(() => messages.id, { onDelete: "cascade" }),
    filename: text("filename").notNull(),
    mimeType: text("mime_type").notNull(),
    objectKey: text("object_key").notNull(),
    sizeBytes: integer("size_bytes").notNull(),
    reportType: text("report_type"),
    version: integer("version").notNull().default(1),
    slideCount: integer("slide_count"),
    presentationJson: text("presentation_json"),
    createdAt: text("created_at").notNull(),
  },
  (table) => [
    index("idx_artifacts_chat_created").on(table.chatId, table.createdAt),
    uniqueIndex("idx_artifacts_message").on(table.messageId),
  ],
);

export const modelUsage = sqliteTable(
  "model_usage",
  {
    id: text("id").primaryKey(),
    projectId: text("project_id"),
    chatId: text("chat_id"),
    requestId: text("request_id").notNull(),
    provider: text("provider").notNull(),
    modelKey: text("model_key").notNull(),
    inputTokens: integer("input_tokens"),
    outputTokens: integer("output_tokens"),
    estimatedCostUsd: integer("estimated_cost_microusd"),
    latencyMs: integer("latency_ms"),
    requestType: text("request_type").notNull(),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("idx_model_usage_chat").on(table.chatId, table.createdAt)],
);
