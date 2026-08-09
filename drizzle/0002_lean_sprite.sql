ALTER TABLE `artifacts` ADD `format` text DEFAULT 'pptx' NOT NULL;--> statement-breakpoint
ALTER TABLE `artifacts` ADD `unit_count` integer;--> statement-breakpoint
ALTER TABLE `artifacts` ADD `unit_label` text;--> statement-breakpoint
ALTER TABLE `artifacts` ADD `model_json` text;--> statement-breakpoint
ALTER TABLE `artifacts` ADD `parent_artifact_id` text;