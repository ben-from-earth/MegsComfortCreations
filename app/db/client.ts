import { drizzle } from 'drizzle-orm/node-postgres';
import { Pool } from 'pg';

const runtimeDatabaseUrl = process.env.DATABASE_URL;

if (!runtimeDatabaseUrl) {
  throw new Error(
    'Missing DATABASE_URL. Set DATABASE_URL for the current environment before starting the app.',
  );
}

const pool = new Pool({ connectionString: runtimeDatabaseUrl });
export const db = drizzle({ client: pool });
export type Db = typeof db;
