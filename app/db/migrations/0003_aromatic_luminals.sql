CREATE TABLE "google_api_query_usage" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"date" text NOT NULL,
	"query_count" integer DEFAULT 0 NOT NULL,
	CONSTRAINT "google_api_query_usage_date_unique" UNIQUE("date")
);
