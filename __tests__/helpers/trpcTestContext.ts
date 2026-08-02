import { appRouter } from 'lib/trpc/routers/_app';
import type { Context } from 'lib/trpc/context';

type TestUser = {
  id: string;
  role: 'admin' | 'user';
};

/**
 * Auth identity + mockable db for router unit tests.
 * Live Context also carries full Better Auth session / Drizzle client shapes.
 */
export type TrpcTestContext = {
  db: object;
  authSession: { user: TestUser } | null;
  user: TestUser | null;
};

export function createAdminContext(db: object = {}): TrpcTestContext {
  const user: TestUser = { id: 'admin-user-id', role: 'admin' };
  return {
    db,
    authSession: { user },
    user,
  };
}

export function createUserContext(db: object = {}): TrpcTestContext {
  const user: TestUser = { id: 'normal-user-id', role: 'user' };
  return {
    db,
    authSession: { user },
    user,
  };
}

export function createAnonContext(db: object = {}): TrpcTestContext {
  return {
    db,
    authSession: null,
    user: null,
  };
}

export function createTrpcCaller(ctx: TrpcTestContext) {
  // Test doubles omit live Better Auth / Drizzle runtime fields on purpose.
  return appRouter.createCaller(ctx as Context);
}
