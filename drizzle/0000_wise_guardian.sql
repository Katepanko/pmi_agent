CREATE TABLE `chats` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`project_id` text,
	`title` text NOT NULL,
	`model_key` text DEFAULT 'openai-gpt56' NOT NULL,
	`audience` text DEFAULT 'Steering Committee' NOT NULL,
	`archived_at` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `idx_chats_user_project` ON `chats` (`user_id`,`project_id`);--> statement-breakpoint
CREATE INDEX `idx_chats_user_updated` ON `chats` (`user_id`,`updated_at`);--> statement-breakpoint
CREATE TABLE `knowledge_versions` (
	`id` text PRIMARY KEY NOT NULL,
	`project_id` text,
	`chat_id` text,
	`version` integer NOT NULL,
	`scope` text NOT NULL,
	`content_json` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`chat_id`) REFERENCES `chats`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE UNIQUE INDEX `idx_knowledge_scope_version` ON `knowledge_versions` (`scope`,`project_id`,`chat_id`,`version`);--> statement-breakpoint
CREATE TABLE `messages` (
	`id` text PRIMARY KEY NOT NULL,
	`chat_id` text NOT NULL,
	`role` text NOT NULL,
	`content` text NOT NULL,
	`classification` text DEFAULT 'conversation' NOT NULL,
	`model_key` text,
	`source_coverage_json` text,
	`stopped_at` text,
	`created_at` text NOT NULL,
	FOREIGN KEY (`chat_id`) REFERENCES `chats`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `idx_messages_chat_created` ON `messages` (`chat_id`,`created_at`);--> statement-breakpoint
CREATE TABLE `model_usage` (
	`id` text PRIMARY KEY NOT NULL,
	`project_id` text,
	`chat_id` text,
	`request_id` text NOT NULL,
	`provider` text NOT NULL,
	`model_key` text NOT NULL,
	`input_tokens` integer,
	`output_tokens` integer,
	`estimated_cost_microusd` integer,
	`latency_ms` integer,
	`request_type` text NOT NULL,
	`created_at` text NOT NULL
);
--> statement-breakpoint
CREATE INDEX `idx_model_usage_chat` ON `model_usage` (`chat_id`,`created_at`);--> statement-breakpoint
CREATE TABLE `projects` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`name` text NOT NULL,
	`description` text DEFAULT '' NOT NULL,
	`companies` text DEFAULT '' NOT NULL,
	`deal_rationale` text DEFAULT '' NOT NULL,
	`integration_type` text DEFAULT '' NOT NULL,
	`day_one_date` text,
	`objectives` text DEFAULT '' NOT NULL,
	`synergy_targets` text DEFAULT '' NOT NULL,
	`governance` text DEFAULT '' NOT NULL,
	`terminology` text DEFAULT '' NOT NULL,
	`reporting_expectations` text DEFAULT '' NOT NULL,
	`instructions` text DEFAULT '' NOT NULL,
	`icon` text DEFAULT 'layers' NOT NULL,
	`archived_at` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL
);
--> statement-breakpoint
CREATE INDEX `idx_projects_user_archived` ON `projects` (`user_id`,`archived_at`);--> statement-breakpoint
CREATE TABLE `report_draft_versions` (
	`id` text PRIMARY KEY NOT NULL,
	`draft_id` text NOT NULL,
	`version` integer NOT NULL,
	`content` text NOT NULL,
	`sections_json` text DEFAULT '[]' NOT NULL,
	`sources_json` text DEFAULT '[]' NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`draft_id`) REFERENCES `report_drafts`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE UNIQUE INDEX `idx_draft_version` ON `report_draft_versions` (`draft_id`,`version`);--> statement-breakpoint
CREATE TABLE `report_drafts` (
	`id` text PRIMARY KEY NOT NULL,
	`project_id` text,
	`chat_id` text NOT NULL,
	`title` text NOT NULL,
	`audience` text NOT NULL,
	`report_type` text NOT NULL,
	`requested_format` text,
	`current_version` integer DEFAULT 1 NOT NULL,
	`based_on_knowledge_version` integer DEFAULT 1 NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`chat_id`) REFERENCES `chats`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `idx_report_drafts_chat` ON `report_drafts` (`chat_id`,`updated_at`);--> statement-breakpoint
CREATE TABLE `source_priority_rules` (
	`id` text PRIMARY KEY NOT NULL,
	`project_id` text,
	`chat_id` text,
	`scope` text NOT NULL,
	`source_id` text NOT NULL,
	`applies_to` text NOT NULL,
	`instruction` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`chat_id`) REFERENCES `chats`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`source_id`) REFERENCES `sources`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `idx_source_rules_chat` ON `source_priority_rules` (`chat_id`);--> statement-breakpoint
CREATE TABLE `source_segments` (
	`id` text PRIMARY KEY NOT NULL,
	`source_id` text NOT NULL,
	`ordinal` integer NOT NULL,
	`location` text NOT NULL,
	`kind` text NOT NULL,
	`content` text NOT NULL,
	`structured_json` text,
	`confidence` text DEFAULT 'high' NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`source_id`) REFERENCES `sources`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE UNIQUE INDEX `idx_source_segments_source_ordinal` ON `source_segments` (`source_id`,`ordinal`);--> statement-breakpoint
CREATE INDEX `idx_source_segments_kind` ON `source_segments` (`kind`);--> statement-breakpoint
CREATE TABLE `sources` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`project_id` text,
	`chat_id` text NOT NULL,
	`message_id` text,
	`file_name` text NOT NULL,
	`file_type` text NOT NULL,
	`media_type` text NOT NULL,
	`size_bytes` integer NOT NULL,
	`object_key` text NOT NULL,
	`checksum` text,
	`extraction_status` text DEFAULT 'pending' NOT NULL,
	`extraction_warnings_json` text DEFAULT '[]' NOT NULL,
	`metadata_json` text DEFAULT '{}' NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`chat_id`) REFERENCES `chats`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`message_id`) REFERENCES `messages`(`id`) ON UPDATE no action ON DELETE set null
);
--> statement-breakpoint
CREATE INDEX `idx_sources_chat_status` ON `sources` (`chat_id`,`extraction_status`);--> statement-breakpoint
CREATE INDEX `idx_sources_project` ON `sources` (`project_id`);