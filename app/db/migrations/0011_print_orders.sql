DO $migrate$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM information_schema.columns
    WHERE table_schema = 'public'
      AND table_name = 'orders'
      AND column_name = 'customer_name'
  ) THEN
    IF NOT EXISTS (
      SELECT 1
      FROM information_schema.columns
      WHERE table_schema = 'public'
        AND table_name = 'orders'
        AND column_name = 'customer_id'
    ) THEN
      RAISE EXCEPTION 'orders has neither customer_name nor customer_id';
    END IF;

    IF EXISTS (SELECT 1 FROM "public"."orders") THEN
      RAISE EXCEPTION 'refusing to replace orders because it contains rows';
    END IF;

    IF to_regclass('public.orders_books') IS NOT NULL
       AND EXISTS (SELECT 1 FROM "public"."orders_books") THEN
      RAISE EXCEPTION 'refusing to replace orders_books because it contains rows';
    END IF;

    IF to_regclass('public.pngs') IS NOT NULL
       AND EXISTS (SELECT 1 FROM "public"."pngs") THEN
      RAISE EXCEPTION 'refusing to replace pngs because it contains rows';
    END IF;

    DROP TABLE IF EXISTS "public"."orders_books";
    DROP TABLE "public"."orders";
    DROP TABLE IF EXISTS "public"."pngs";

    CREATE TABLE "public"."orders" (
      "id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
      "customer_name" text NOT NULL,
      "order_number" text NOT NULL,
      "png_format" integer NOT NULL,
      "repeat_count" integer NOT NULL,
      "file_url" text NOT NULL,
      "file_name" text NOT NULL,
      "mime_type" text NOT NULL,
      "created_at" timestamp DEFAULT now() NOT NULL,
      "updated_at" timestamp DEFAULT now() NOT NULL,
      CONSTRAINT "orders_order_number_unique" UNIQUE("order_number"),
      CONSTRAINT "orders_png_format_check" CHECK ("png_format" in (3, 5)),
      CONSTRAINT "orders_repeat_count_check" CHECK ("repeat_count" between 1 and 25)
    );
  END IF;

  IF to_regclass('public.order_covers') IS NULL THEN
    CREATE TABLE "public"."order_covers" (
      "id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
      "order_id" uuid NOT NULL,
      "media_image_item_id" uuid NOT NULL,
      "position" integer NOT NULL,
      CONSTRAINT "order_covers_order_position_unique" UNIQUE("order_id","position"),
      CONSTRAINT "order_covers_position_check" CHECK ("position" >= 0)
    );

    ALTER TABLE "public"."order_covers"
      ADD CONSTRAINT "order_covers_order_id_orders_id_fk"
      FOREIGN KEY ("order_id") REFERENCES "public"."orders"("id")
      ON DELETE cascade ON UPDATE no action;

    ALTER TABLE "public"."order_covers"
      ADD CONSTRAINT "order_covers_media_image_item_id_media_image_items_id_fk"
      FOREIGN KEY ("media_image_item_id") REFERENCES "public"."media_image_items"("id")
      ON DELETE restrict ON UPDATE no action;

    CREATE INDEX "order_covers_media_image_item_id_idx"
      ON "public"."order_covers" USING btree ("media_image_item_id");
  END IF;
END
$migrate$;
