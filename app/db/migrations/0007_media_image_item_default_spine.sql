ALTER TABLE "media_image_items" ADD COLUMN "is_default" boolean DEFAULT false NOT NULL;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD COLUMN "spine_color" text DEFAULT '#ffffff' NOT NULL;
--> statement-breakpoint
UPDATE "media_image_items" AS mii
SET "spine_color" = b."spine_color"
FROM "books" AS b
WHERE mii."book_id" = b."id";
--> statement-breakpoint
UPDATE "media_image_items" AS mii
SET "spine_color" = om."spine_color"
FROM "other_media" AS om
WHERE mii."other_media_id" = om."id";
--> statement-breakpoint
WITH ranked_book_images AS (
  SELECT
    "id",
    row_number() OVER (
      PARTITION BY "book_id"
      ORDER BY "sort_order" ASC, "created_at" ASC, "id" ASC
    ) AS image_rank
  FROM "media_image_items"
  WHERE "book_id" IS NOT NULL
)
UPDATE "media_image_items" AS mii
SET "is_default" = (ranked_book_images."image_rank" = 1)
FROM ranked_book_images
WHERE ranked_book_images."id" = mii."id";
--> statement-breakpoint
WITH ranked_other_media_images AS (
  SELECT
    "id",
    row_number() OVER (
      PARTITION BY "other_media_id"
      ORDER BY "sort_order" ASC, "created_at" ASC, "id" ASC
    ) AS image_rank
  FROM "media_image_items"
  WHERE "other_media_id" IS NOT NULL
)
UPDATE "media_image_items" AS mii
SET "is_default" = (ranked_other_media_images."image_rank" = 1)
FROM ranked_other_media_images
WHERE ranked_other_media_images."id" = mii."id";
--> statement-breakpoint
CREATE UNIQUE INDEX "media_image_items_book_default_unique"
ON "media_image_items" USING btree ("book_id")
WHERE "book_id" IS NOT NULL AND "is_default" = true;
--> statement-breakpoint
CREATE UNIQUE INDEX "media_image_items_other_media_default_unique"
ON "media_image_items" USING btree ("other_media_id")
WHERE "other_media_id" IS NOT NULL AND "is_default" = true;
--> statement-breakpoint
CREATE INDEX "media_image_items_book_is_default_idx"
ON "media_image_items" USING btree ("book_id", "is_default");
--> statement-breakpoint
CREATE INDEX "media_image_items_other_media_is_default_idx"
ON "media_image_items" USING btree ("other_media_id", "is_default");
