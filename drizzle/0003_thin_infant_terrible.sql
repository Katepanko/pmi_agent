ALTER TABLE `artifacts` ADD `content_language` text DEFAULT 'und' NOT NULL;--> statement-breakpoint
ALTER TABLE `artifacts` ADD `template_source_id` text;--> statement-breakpoint
ALTER TABLE `artifacts` ADD `generation_operation` text DEFAULT 'generate' NOT NULL;