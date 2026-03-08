import { db } from '@/db/client';
import { auth } from 'lib/auth';

type AuthSession = Awaited<ReturnType<typeof auth.api.getSession>>;
type AuthenticatedSession = NonNullable<AuthSession>;
type SessionUser = AuthenticatedSession extends { user: infer User }
  ? User
  : never;

export type Context = {
  db: typeof db;
  authSession: AuthSession | null;
  user: SessionUser | null;
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
