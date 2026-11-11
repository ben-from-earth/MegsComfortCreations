ALTER TABLE "media_image_items" ADD COLUMN "is_default" boolean DEFAULT false NOT NULL;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD COLUMN "spine_color" text DEFAULT '#ffffff' NOT NULL;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD COLUMN "code_color" text;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_code_color_check" CHECK (
	("album_id" IS NOT NULL) OR ("code_color" IS NULL)
);
--> statement-breakpoint
UPDATE "media_image_items" AS mii
SET "spine_color" = b."spine_color"
FROM "books" AS b
WHERE mii."book_id" = b."id";
--> statement-breakpoint
UPDATE "media_image_items" AS mii
SET "spine_color" = m."spine_color"
FROM "movies" AS m
WHERE mii."movie_id" = m."id";
--> statement-breakpoint
UPDATE "media_image_items" AS mii
SET "spine_color" = vg."spine_color"
FROM "video_games" AS vg
WHERE mii."video_game_id" = vg."id";
--> statement-breakpoint
UPDATE "media_image_items" AS mii
SET "spine_color" = a."spine_color"
FROM "albums" AS a
WHERE mii."album_id" = a."id";
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
WITH ranked_movie_images AS (
  SELECT
    "id",
    row_number() OVER (
      PARTITION BY "movie_id"
      ORDER BY "sort_order" ASC, "created_at" ASC, "id" ASC
    ) AS image_rank
  FROM "media_image_items"
  WHERE "movie_id" IS NOT NULL
)
UPDATE "media_image_items" AS mii
SET "is_default" = (ranked_movie_images."image_rank" = 1)
FROM ranked_movie_images
WHERE ranked_movie_images."id" = mii."id";
--> statement-breakpoint
WITH ranked_video_game_images AS (
  SELECT
    "id",
    row_number() OVER (
      PARTITION BY "video_game_id"
      ORDER BY "sort_order" ASC, "created_at" ASC, "id" ASC
    ) AS image_rank
  FROM "media_image_items"
  WHERE "video_game_id" IS NOT NULL
)
UPDATE "media_image_items" AS mii
SET "is_default" = (ranked_video_game_images."image_rank" = 1)
FROM ranked_video_game_images
WHERE ranked_video_game_images."id" = mii."id";
--> statement-breakpoint
WITH ranked_album_images AS (
  SELECT
    "id",
    row_number() OVER (
      PARTITION BY "album_id"
      ORDER BY "sort_order" ASC, "created_at" ASC, "id" ASC
    ) AS image_rank
  FROM "media_image_items"
  WHERE "album_id" IS NOT NULL
)
UPDATE "media_image_items" AS mii
SET "is_default" = (ranked_album_images."image_rank" = 1)
FROM ranked_album_images
WHERE ranked_album_images."id" = mii."id";
--> statement-breakpoint
CREATE UNIQUE INDEX "media_image_items_book_default_unique"
ON "media_image_items" USING btree ("book_id")
WHERE "book_id" IS NOT NULL AND "is_default" = true;
--> statement-breakpoint
CREATE UNIQUE INDEX "media_image_items_movie_default_unique"
ON "media_image_items" USING btree ("movie_id")
WHERE "movie_id" IS NOT NULL AND "is_default" = true;
--> statement-breakpoint
CREATE UNIQUE INDEX "media_image_items_video_game_default_unique"
ON "media_image_items" USING btree ("video_game_id")
WHERE "video_game_id" IS NOT NULL AND "is_default" = true;
--> statement-breakpoint
CREATE UNIQUE INDEX "media_image_items_album_default_unique"
ON "media_image_items" USING btree ("album_id")
WHERE "album_id" IS NOT NULL AND "is_default" = true;
--> statement-breakpoint
CREATE INDEX "media_image_items_book_is_default_idx"
ON "media_image_items" USING btree ("book_id", "is_default");
--> statement-breakpoint
CREATE INDEX "media_image_items_movie_is_default_idx"
ON "media_image_items" USING btree ("movie_id", "is_default");
--> statement-breakpoint
CREATE INDEX "media_image_items_video_game_is_default_idx"
ON "media_image_items" USING btree ("video_game_id", "is_default");
--> statement-breakpoint
CREATE INDEX "media_image_items_album_is_default_idx"
ON "media_image_items" USING btree ("album_id", "is_default");
