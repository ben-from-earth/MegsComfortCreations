CREATE TYPE "public"."other_media_type" AS ENUM('movie', 'videoGame', 'album');--> statement-breakpoint
CREATE TYPE "public"."movie_format" AS ENUM('dvd', 'vhs', 'bluRay');--> statement-breakpoint
CREATE TABLE "other_media" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"media_type" "other_media_type" NOT NULL,
	"title" text NOT NULL,
	"spine_color" text NOT NULL,
	"movie_format" "movie_format",
	"image_urls" text[] NOT NULL,
	CONSTRAINT "other_media_movie_format_check" CHECK (
		("media_type" = 'movie' AND "movie_format" IS NOT NULL)
		OR ("media_type" <> 'movie' AND "movie_format" IS NULL)
	)
);
--> statement-breakpoint
CREATE UNIQUE INDEX "other_media_type_title_unique" ON "other_media" USING btree ("media_type","title") WHERE "movie_format" IS NULL;
--> statement-breakpoint
CREATE UNIQUE INDEX "other_media_movie_title_format_unique" ON "other_media" USING btree ("media_type","title","movie_format") WHERE "movie_format" IS NOT NULL;
--> statement-breakpoint
CREATE INDEX "other_media_media_type_idx" ON "other_media" USING btree ("media_type");
--> statement-breakpoint
CREATE INDEX "other_media_title_idx" ON "other_media" USING btree ("title");
--> statement-breakpoint
INSERT INTO "other_media" ("media_type", "title", "spine_color", "image_urls", "movie_format")
SELECT 'movie', "title", "spine_color", "image_urls", 'dvd'
FROM "movies";
--> statement-breakpoint
INSERT INTO "other_media" ("media_type", "title", "spine_color", "image_urls")
SELECT 'videoGame', "title", "spine_color", "image_urls"
FROM "video_games";
--> statement-breakpoint
INSERT INTO "other_media" ("media_type", "title", "spine_color", "image_urls")
SELECT 'album', "title", "spine_color", "image_urls"
FROM "albums";
--> statement-breakpoint
DROP TABLE IF EXISTS "movies", "video_games", "albums";
