DELETE FROM "public"."movies" WHERE "movie_format" = 'bluRay';
--> statement-breakpoint
ALTER TYPE "public"."movie_format" RENAME TO "movie_format_old";
--> statement-breakpoint
CREATE TYPE "public"."movie_format" AS ENUM('dvd', 'vhs');
--> statement-breakpoint
ALTER TABLE "public"."movies" ALTER COLUMN "movie_format" SET DATA TYPE "public"."movie_format" USING "movie_format"::text::"public"."movie_format";
--> statement-breakpoint
DROP TYPE "public"."movie_format_old";
