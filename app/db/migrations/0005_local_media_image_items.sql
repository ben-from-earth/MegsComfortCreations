CREATE TABLE "media_image_items" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"book_id" uuid,
	"movie_id" uuid,
	"video_game_id" uuid,
	"album_id" uuid,
	"path" text NOT NULL,
	"source_url" text,
	"mime_type" text,
	"size_bytes" integer,
	"sort_order" integer DEFAULT 0 NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL,
	CONSTRAINT "media_image_items_parent_check" CHECK ((
		("book_id" IS NOT NULL)::int +
		("movie_id" IS NOT NULL)::int +
		("video_game_id" IS NOT NULL)::int +
		("album_id" IS NOT NULL)::int
	) = 1)
);
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_book_id_books_id_fk" FOREIGN KEY ("book_id") REFERENCES "public"."books"("id") ON DELETE cascade ON UPDATE no action;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_movie_id_movies_id_fk" FOREIGN KEY ("movie_id") REFERENCES "public"."movies"("id") ON DELETE cascade ON UPDATE no action;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_video_game_id_video_games_id_fk" FOREIGN KEY ("video_game_id") REFERENCES "public"."video_games"("id") ON DELETE cascade ON UPDATE no action;
--> statement-breakpoint
ALTER TABLE "media_image_items" ADD CONSTRAINT "media_image_items_album_id_albums_id_fk" FOREIGN KEY ("album_id") REFERENCES "public"."albums"("id") ON DELETE cascade ON UPDATE no action;
--> statement-breakpoint
CREATE INDEX "media_image_items_book_id_idx" ON "media_image_items" USING btree ("book_id");
--> statement-breakpoint
CREATE INDEX "media_image_items_movie_id_idx" ON "media_image_items" USING btree ("movie_id");
--> statement-breakpoint
CREATE INDEX "media_image_items_video_game_id_idx" ON "media_image_items" USING btree ("video_game_id");
--> statement-breakpoint
CREATE INDEX "media_image_items_album_id_idx" ON "media_image_items" USING btree ("album_id");
--> statement-breakpoint
CREATE INDEX "media_image_items_book_sort_order_idx" ON "media_image_items" USING btree ("book_id","sort_order");
--> statement-breakpoint
CREATE INDEX "media_image_items_movie_sort_order_idx" ON "media_image_items" USING btree ("movie_id","sort_order");
--> statement-breakpoint
CREATE INDEX "media_image_items_video_game_sort_order_idx" ON "media_image_items" USING btree ("video_game_id","sort_order");
--> statement-breakpoint
CREATE INDEX "media_image_items_album_sort_order_idx" ON "media_image_items" USING btree ("album_id","sort_order");
