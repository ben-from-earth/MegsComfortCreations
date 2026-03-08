import { drizzle } from 'drizzle-orm/neon-http';
import { neon } from '@neondatabase/serverless';

const runtimeDatabaseUrl = process.env.DATABASE_URL;

if (!runtimeDatabaseUrl) {
  throw new Error(
    'Missing DATABASE_URL. Set DATABASE_URL for the current environment before starting the app.',
  );
}

const sql = neon(runtimeDatabaseUrl);
export const db = drizzle({ client: sql });
export type Db = typeof db;
