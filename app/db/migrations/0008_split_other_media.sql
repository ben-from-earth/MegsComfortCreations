DO $migrate$
DECLARE
  index_name text;
BEGIN
  IF to_regclass('public.other_media') IS NULL THEN
    RETURN;
  END IF;

  IF to_regclass('public.movies') IS NOT NULL
    OR to_regclass('public.video_games') IS NOT NULL
    OR to_regclass('public.albums') IS NOT NULL THEN
    RAISE EXCEPTION 'other_media cannot be split because movies, video_games, or albums already exist';
  END IF;

  CREATE TYPE "public"."movie_format" AS ENUM('dvd', 'vhs', 'bluRay');

  CREATE TABLE "public"."movies" (
    "id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
    "title" text NOT NULL,
    "year" integer,
    "spine_color" text NOT NULL,
    "movie_format" "public"."movie_format" NOT NULL,
    CONSTRAINT "movies_title_format_unique" UNIQUE NULLS NOT DISTINCT ("title", "year", "movie_format")
  );

  CREATE TABLE "public"."video_games" (
    "id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
    "title" text NOT NULL,
    "year" integer,
    "console" text DEFAULT '' NOT NULL,
    "spine_color" text NOT NULL,
    CONSTRAINT "video_games_title_unique" UNIQUE NULLS NOT DISTINCT ("title", "year", "console")
  );

  CREATE TABLE "public"."albums" (
    "id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
    "title" text NOT NULL,
    "artist" text NOT NULL,
    "release_date" date,
    "spotify_album_uri" text,
    "spine_color" text NOT NULL
  );
  CREATE UNIQUE INDEX "albums_title_artist_unique"
    ON "public"."albums" USING btree ("title", "artist");
  CREATE UNIQUE INDEX "albums_spotify_album_uri_unique"
    ON "public"."albums" USING btree ("spotify_album_uri")
    WHERE "spotify_album_uri" IS NOT NULL;

  IF EXISTS (
    SELECT 1
    FROM information_schema.columns
    WHERE table_schema = 'public'
      AND table_name = 'other_media'
      AND column_name = 'movie_format'
  ) THEN
    EXECUTE $insert_movies$
      INSERT INTO "public"."movies" ("id", "title", "year", "spine_color", "movie_format")
      SELECT
        "id",
        CASE
          WHEN length(btrim("title")) >= 7
            AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
            AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> ''
          THEN btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6))
          ELSE btrim("title")
        END,
        CASE
          WHEN length(btrim("title")) >= 7
            AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
            AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> ''
          THEN substring(right(btrim("title"), 6) FROM 2 FOR 4)::integer
          ELSE NULL
        END,
        "spine_color",
        COALESCE("movie_format"::text, 'dvd')::"public"."movie_format"
      FROM "public"."other_media"
      WHERE "media_type" = 'movie'
    $insert_movies$;
  ELSE
    INSERT INTO "public"."movies" ("id", "title", "year", "spine_color", "movie_format")
    SELECT
      "id",
      CASE
        WHEN length(btrim("title")) >= 7
          AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
          AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> ''
        THEN btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6))
        ELSE btrim("title")
      END,
      CASE
        WHEN length(btrim("title")) >= 7
          AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
          AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> ''
        THEN substring(right(btrim("title"), 6) FROM 2 FOR 4)::integer
        ELSE NULL
      END,
      "spine_color",
      'dvd'::"public"."movie_format"
    FROM "public"."other_media"
    WHERE "media_type" = 'movie';
  END IF;

  INSERT INTO "public"."video_games" ("id", "title", "year", "console", "spine_color")
  SELECT
    "id",
    CASE
      WHEN length(btrim("title")) >= 7
        AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
        AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> ''
      THEN btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6))
      ELSE btrim("title")
    END,
    CASE
      WHEN length(btrim("title")) >= 7
        AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
        AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> ''
      THEN substring(right(btrim("title"), 6) FROM 2 FOR 4)::integer
      ELSE NULL
    END,
    '',
    "spine_color"
  FROM "public"."other_media"
  WHERE "media_type" = 'videoGame';

  INSERT INTO "public"."albums" (
    "id",
    "title",
    "artist",
    "release_date",
    "spotify_album_uri",
    "spine_color"
  )
  SELECT "id", "title", '', NULL, NULL, "spine_color"
  FROM "public"."other_media"
  WHERE "media_type" = 'album';

  ALTER TABLE "public"."media_image_items"
    ADD COLUMN "movie_id" uuid,
    ADD COLUMN "video_game_id" uuid,
    ADD COLUMN "album_id" uuid,
    ADD COLUMN "code_color" text;

  UPDATE "public"."media_image_items" AS image_item
  SET "movie_id" = media_item."id"
  FROM "public"."other_media" AS media_item
  WHERE image_item."other_media_id" = media_item."id"
    AND media_item."media_type" = 'movie';

  UPDATE "public"."media_image_items" AS image_item
  SET "video_game_id" = media_item."id"
  FROM "public"."other_media" AS media_item
  WHERE image_item."other_media_id" = media_item."id"
    AND media_item."media_type" = 'videoGame';

  UPDATE "public"."media_image_items" AS image_item
  SET "album_id" = media_item."id"
  FROM "public"."other_media" AS media_item
  WHERE image_item."other_media_id" = media_item."id"
    AND media_item."media_type" = 'album';

  IF EXISTS (
    SELECT 1
    FROM "public"."media_image_items"
    WHERE (
      ("book_id" IS NOT NULL)::int +
      ("movie_id" IS NOT NULL)::int +
      ("video_game_id" IS NOT NULL)::int +
      ("album_id" IS NOT NULL)::int
    ) <> 1
  ) THEN
    RAISE EXCEPTION 'media image is missing exactly one parent after leaving other_media';
  END IF;

  ALTER TABLE "public"."media_image_items"
    DROP CONSTRAINT IF EXISTS "media_image_items_parent_check";
  ALTER TABLE "public"."media_image_items"
    DROP CONSTRAINT IF EXISTS "media_image_items_other_media_id_other_media_id_fk";

  FOR index_name IN
    SELECT indexname
    FROM pg_indexes
    WHERE schemaname = 'public'
      AND tablename = 'media_image_items'
      AND indexdef ILIKE '%other_media_id%'
  LOOP
    EXECUTE format('DROP INDEX IF EXISTS "public".%I', index_name);
  END LOOP;

  ALTER TABLE "public"."media_image_items" DROP COLUMN "other_media_id";

  ALTER TABLE "public"."media_image_items"
    ADD CONSTRAINT "media_image_items_movie_id_movies_id_fk"
    FOREIGN KEY ("movie_id") REFERENCES "public"."movies"("id")
    ON DELETE cascade ON UPDATE no action;
  ALTER TABLE "public"."media_image_items"
    ADD CONSTRAINT "media_image_items_video_game_id_video_games_id_fk"
    FOREIGN KEY ("video_game_id") REFERENCES "public"."video_games"("id")
    ON DELETE cascade ON UPDATE no action;
  ALTER TABLE "public"."media_image_items"
    ADD CONSTRAINT "media_image_items_album_id_albums_id_fk"
    FOREIGN KEY ("album_id") REFERENCES "public"."albums"("id")
    ON DELETE cascade ON UPDATE no action;
  ALTER TABLE "public"."media_image_items"
    ADD CONSTRAINT "media_image_items_parent_check" CHECK ((
      ("book_id" IS NOT NULL)::int +
      ("movie_id" IS NOT NULL)::int +
      ("video_game_id" IS NOT NULL)::int +
      ("album_id" IS NOT NULL)::int
    ) = 1);
  ALTER TABLE "public"."media_image_items"
    ADD CONSTRAINT "media_image_items_code_color_check" CHECK (
      ("album_id" IS NOT NULL) OR ("code_color" IS NULL)
    );

  CREATE INDEX "media_image_items_movie_id_idx"
    ON "public"."media_image_items" USING btree ("movie_id");
  CREATE INDEX "media_image_items_video_game_id_idx"
    ON "public"."media_image_items" USING btree ("video_game_id");
  CREATE INDEX "media_image_items_album_id_idx"
    ON "public"."media_image_items" USING btree ("album_id");
  CREATE INDEX "media_image_items_movie_sort_order_idx"
    ON "public"."media_image_items" USING btree ("movie_id", "sort_order");
  CREATE INDEX "media_image_items_video_game_sort_order_idx"
    ON "public"."media_image_items" USING btree ("video_game_id", "sort_order");
  CREATE INDEX "media_image_items_album_sort_order_idx"
    ON "public"."media_image_items" USING btree ("album_id", "sort_order");
  CREATE INDEX "media_image_items_movie_is_default_idx"
    ON "public"."media_image_items" USING btree ("movie_id", "is_default");
  CREATE INDEX "media_image_items_video_game_is_default_idx"
    ON "public"."media_image_items" USING btree ("video_game_id", "is_default");
  CREATE INDEX "media_image_items_album_is_default_idx"
    ON "public"."media_image_items" USING btree ("album_id", "is_default");
  CREATE UNIQUE INDEX "media_image_items_movie_default_unique"
    ON "public"."media_image_items" USING btree ("movie_id")
    WHERE "movie_id" IS NOT NULL AND "is_default" = true;
  CREATE UNIQUE INDEX "media_image_items_video_game_default_unique"
    ON "public"."media_image_items" USING btree ("video_game_id")
    WHERE "video_game_id" IS NOT NULL AND "is_default" = true;
  CREATE UNIQUE INDEX "media_image_items_album_default_unique"
    ON "public"."media_image_items" USING btree ("album_id")
    WHERE "album_id" IS NOT NULL AND "is_default" = true;

  DROP TABLE "public"."other_media";
  DROP TYPE "public"."other_media_type";
END
$migrate$;
