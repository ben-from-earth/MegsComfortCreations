CREATE TYPE "public"."movie_format" AS ENUM('dvd', 'vhs', 'bluRay');
--> statement-breakpoint
ALTER TABLE "movies" ADD COLUMN "movie_format" "movie_format" DEFAULT 'dvd' NOT NULL;
--> statement-breakpoint
ALTER TABLE "movies" ALTER COLUMN "movie_format" DROP DEFAULT;
--> statement-breakpoint
ALTER TABLE "movies" DROP CONSTRAINT "movies_title_unique";
--> statement-breakpoint
CREATE UNIQUE INDEX "movies_title_format_unique" ON "movies" USING btree ("title","movie_format");
--> statement-breakpoint
ALTER TABLE "albums" ADD COLUMN "artist" text DEFAULT '' NOT NULL;
--> statement-breakpoint
ALTER TABLE "albums" ALTER COLUMN "artist" DROP DEFAULT;
--> statement-breakpoint
ALTER TABLE "albums" ADD COLUMN "release_date" date;
--> statement-breakpoint
ALTER TABLE "albums" ADD COLUMN "spotify_album_uri" text;
--> statement-breakpoint
ALTER TABLE "albums" DROP CONSTRAINT "albums_title_unique";
--> statement-breakpoint
CREATE UNIQUE INDEX "albums_title_artist_unique" ON "albums" USING btree ("title","artist");
--> statement-breakpoint
CREATE UNIQUE INDEX "albums_spotify_album_uri_unique" ON "albums" USING btree ("spotify_album_uri") WHERE "spotify_album_uri" IS NOT NULL;
