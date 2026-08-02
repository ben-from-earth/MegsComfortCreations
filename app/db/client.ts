import { ExtractTablesWithRelations } from 'drizzle-orm';
import { NodePgQueryResultHKT, drizzle } from 'drizzle-orm/node-postgres';
import { PgTransaction } from 'drizzle-orm/pg-core';
import { Pool } from 'pg';
import * as schema from './schema';

const runtimeDatabaseUrl = process.env.DATABASE_URL;

if (!runtimeDatabaseUrl) {
  throw new Error(
    'Missing DATABASE_URL. Set DATABASE_URL for the current environment before starting the app.',
  );
}

const pool = new Pool({ connectionString: runtimeDatabaseUrl });
export const db = drizzle({ client: pool, schema });
export type Db = typeof db;
export type DBSchema = typeof schema;
export type DbTransaction = PgTransaction<
  NodePgQueryResultHKT,
  DBSchema,
  ExtractTablesWithRelations<DBSchema>
>;
export type DbExecutor = Db | DbTransaction;
