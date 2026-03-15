CREATE TABLE "media_image_items" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"book_id" uuid,
	"other_media_id" uuid,
	"path" text NOT NULL,
	"source_url" text,
	"mime_type" text,
	"size_bytes" integer,
	"sort_order" integer DEFAULT 0 NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL,
	CONSTRAINT "media_image_items_parent_check" CHECK (
		("book_id" IS NOT NULL AND "other_media_id" IS NULL)
		OR ("book_id" IS NULL AND "other_media_id" IS NOT NULL)
	)
);
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_book_id_books_id_fk" FOREIGN KEY ("book_id") REFERENCES "public"."books"("id") ON DELETE cascade ON UPDATE no action;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_other_media_id_other_media_id_fk" FOREIGN KEY ("other_media_id") REFERENCES "public"."other_media"("id") ON DELETE cascade ON UPDATE no action;
--> statement-breakpoint
CREATE INDEX "media_image_items_book_id_idx" ON "media_image_items" USING btree ("book_id");
--> statement-breakpoint
CREATE INDEX "media_image_items_other_media_id_idx" ON "media_image_items" USING btree ("other_media_id");
--> statement-breakpoint
CREATE INDEX "media_image_items_book_sort_order_idx" ON "media_image_items" USING btree ("book_id","sort_order");
--> statement-breakpoint
CREATE INDEX "media_image_items_other_media_sort_order_idx" ON "media_image_items" USING btree ("other_media_id","sort_order");
