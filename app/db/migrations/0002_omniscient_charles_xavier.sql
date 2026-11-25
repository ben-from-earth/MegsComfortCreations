CREATE TABLE "customers" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"first_name" text NOT NULL,
	"last_name" text NOT NULL,
	"address_line_1" text NOT NULL,
	"address_line_2" text,
	"city" text NOT NULL,
	"state" text NOT NULL,
	"postal_code" text NOT NULL,
	"country" text NOT NULL,
	"phone" text,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "customers_users" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"customer_id" uuid NOT NULL,
	"user_id" uuid NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "orders" (
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
--> statement-breakpoint
ALTER TABLE "albums" RENAME COLUMN "spineColor" TO "spine_color";--> statement-breakpoint
ALTER TABLE "albums" RENAME COLUMN "imageUrls" TO "image_urls";--> statement-breakpoint
ALTER TABLE "books" RENAME COLUMN "pageCount" TO "page_count";--> statement-breakpoint
ALTER TABLE "books" RENAME COLUMN "pubYear" TO "pub_year";--> statement-breakpoint
ALTER TABLE "books" RENAME COLUMN "spineColor" TO "spine_color";--> statement-breakpoint
ALTER TABLE "books" RENAME COLUMN "imageUrls" TO "image_urls";--> statement-breakpoint
ALTER TABLE "movies" RENAME COLUMN "spineColor" TO "spine_color";--> statement-breakpoint
ALTER TABLE "movies" RENAME COLUMN "imageUrls" TO "image_urls";--> statement-breakpoint
ALTER TABLE "video_games" RENAME COLUMN "spineColor" TO "spine_color";--> statement-breakpoint
ALTER TABLE "video_games" RENAME COLUMN "imageUrls" TO "image_urls";--> statement-breakpoint
ALTER TABLE "customers_users" ADD CONSTRAINT "customers_users_customer_id_customers_id_fk" FOREIGN KEY ("customer_id") REFERENCES "public"."customers"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "customers_users" ADD CONSTRAINT "customers_users_user_id_users_id_fk" FOREIGN KEY ("user_id") REFERENCES "public"."users"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
CREATE UNIQUE INDEX "customers_users_customer_user_unique" ON "customers_users" USING btree ("customer_id","user_id");