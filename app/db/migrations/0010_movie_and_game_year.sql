DO $migrate$
DECLARE
  movie_columns text[];
  movie_nulls_not_distinct boolean;
  game_columns text[];
  game_nulls_not_distinct boolean;
BEGIN
  ALTER TABLE "public"."movies" ADD COLUMN IF NOT EXISTS "year" integer;

  ALTER TABLE "public"."video_games" ADD COLUMN IF NOT EXISTS "year" integer;
  ALTER TABLE "public"."video_games"
    ADD COLUMN IF NOT EXISTS "console" text DEFAULT '' NOT NULL;

  UPDATE "public"."video_games"
  SET "console" = ''
  WHERE "console" IS NULL;

  ALTER TABLE "public"."video_games" ALTER COLUMN "console" SET DEFAULT '';
  ALTER TABLE "public"."video_games" ALTER COLUMN "console" SET NOT NULL;

  UPDATE "public"."movies"
  SET
    "year" = COALESCE(
      "year",
      substring(right(btrim("title"), 6) FROM 2 FOR 4)::integer
    ),
    "title" = btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6))
  WHERE length(btrim("title")) >= 7
    AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
    AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> '';

  UPDATE "public"."video_games"
  SET
    "year" = COALESCE(
      "year",
      substring(right(btrim("title"), 6) FROM 2 FOR 4)::integer
    ),
    "title" = btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6))
  WHERE length(btrim("title")) >= 7
    AND right(btrim("title"), 6) ~ '^\([0-9]{4}\)$'
    AND btrim(substring(btrim("title") FROM 1 FOR length(btrim("title")) - 6)) <> '';

  SELECT
    array_agg(attribute.attname ORDER BY key.ordinality),
    index.indnullsnotdistinct
  INTO movie_columns, movie_nulls_not_distinct
  FROM pg_index AS index
  JOIN pg_class AS index_class ON index_class.oid = index.indexrelid
  JOIN pg_namespace AS index_namespace ON index_namespace.oid = index_class.relnamespace
  JOIN LATERAL unnest(index.indkey) WITH ORDINALITY AS key(attnum, ordinality) ON true
  JOIN pg_attribute AS attribute
    ON attribute.attrelid = index.indrelid
   AND attribute.attnum = key.attnum
  WHERE index_namespace.nspname = 'public'
    AND index_class.relname = 'movies_title_format_unique'
  GROUP BY index.indnullsnotdistinct;

  IF movie_columns IS DISTINCT FROM ARRAY['title', 'year', 'movie_format']::text[]
     OR movie_nulls_not_distinct IS DISTINCT FROM true THEN
    IF EXISTS (
      SELECT 1
      FROM pg_constraint
      WHERE conname = 'movies_title_format_unique'
        AND connamespace = 'public'::regnamespace
    ) THEN
      ALTER TABLE "public"."movies" DROP CONSTRAINT "movies_title_format_unique";
    ELSE
      DROP INDEX IF EXISTS "public"."movies_title_format_unique";
    END IF;

    ALTER TABLE "public"."movies"
      ADD CONSTRAINT "movies_title_format_unique"
      UNIQUE NULLS NOT DISTINCT ("title", "year", "movie_format");
  END IF;

  SELECT
    array_agg(attribute.attname ORDER BY key.ordinality),
    index.indnullsnotdistinct
  INTO game_columns, game_nulls_not_distinct
  FROM pg_index AS index
  JOIN pg_class AS index_class ON index_class.oid = index.indexrelid
  JOIN pg_namespace AS index_namespace ON index_namespace.oid = index_class.relnamespace
  JOIN LATERAL unnest(index.indkey) WITH ORDINALITY AS key(attnum, ordinality) ON true
  JOIN pg_attribute AS attribute
    ON attribute.attrelid = index.indrelid
   AND attribute.attnum = key.attnum
  WHERE index_namespace.nspname = 'public'
    AND index_class.relname = 'video_games_title_unique'
  GROUP BY index.indnullsnotdistinct;

  IF game_columns IS DISTINCT FROM ARRAY['title', 'year', 'console']::text[]
     OR game_nulls_not_distinct IS DISTINCT FROM true THEN
    IF EXISTS (
      SELECT 1
      FROM pg_constraint
      WHERE conname = 'video_games_title_unique'
        AND connamespace = 'public'::regnamespace
    ) THEN
      ALTER TABLE "public"."video_games" DROP CONSTRAINT "video_games_title_unique";
    ELSE
      DROP INDEX IF EXISTS "public"."video_games_title_unique";
    END IF;

    ALTER TABLE "public"."video_games"
      ADD CONSTRAINT "video_games_title_unique"
      UNIQUE NULLS NOT DISTINCT ("title", "year", "console");
  END IF;
END
$migrate$;
