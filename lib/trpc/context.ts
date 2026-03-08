import { db } from '@/db/client';
import { auth } from 'lib/auth';

type AuthSession = Awaited<ReturnType<typeof auth.api.getSession>>;

export type Context = {
  db: typeof db;
  authSession: AuthSession | null;
  user: AuthSession extends { user: infer U } ? U | null : unknown;
};

export async function createContext(opts?: { headers?: Headers }): Promise<Context> {
  let authSession: AuthSession | null = null;

  if (opts?.headers) {
    authSession = await auth.api.getSession({ headers: opts.headers });
  }

  return {
    db,
    authSession,
    user: authSession?.user ?? null,
  };
}
