import { db } from '@//db/client';

export type Context = {
  db: typeof db;
  // Add auth/session when ready, e.g. user?: { id: string };
};

export async function createContext(): Promise<Context> {
  return { db };
}
