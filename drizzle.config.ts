import { config } from 'dotenv';
import { defineConfig } from 'drizzle-kit';

config({ path: '.env' });

const migrationDatabaseUrl =
  process.env.DRIZZLE_DATABASE_URL ?? process.env.DATABASE_URL;

if (!migrationDatabaseUrl) {
  throw new Error(
    'Missing database URL for Drizzle. Set DRIZZLE_DATABASE_URL (preferred) or DATABASE_URL before running migration commands.',
  );
}

export default defineConfig({
  schema: './app/db/schema.ts',
  out: './app/db/migrations',
  dialect: 'postgresql',
  dbCredentials: {
    url: migrationDatabaseUrl,
  },
  casing: 'snake_case',
  verbose: true,
  strict: true,
});
