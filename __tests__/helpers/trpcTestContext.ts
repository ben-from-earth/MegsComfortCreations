import { appRouter } from 'lib/trpc/routers/_app';

type MinimalUser = {
  id: string;
  role: 'admin' | 'user';
};

type MinimalContext = {
  db?: unknown;
  authSession?: { user: MinimalUser } | null;
  user?: MinimalUser | null;
};

export function createAdminContext(db?: unknown): Required<MinimalContext> {
  const user: MinimalUser = { id: 'admin-user-id', role: 'admin' };
  return {
    db: db ?? {},
    authSession: { user },
    user,
  };
}

export function createUserContext(db?: unknown): Required<MinimalContext> {
  const user: MinimalUser = { id: 'normal-user-id', role: 'user' };
  return {
    db: db ?? {},
    authSession: { user },
    user,
  };
}

export function createAnonContext(db?: unknown): Required<MinimalContext> {
  return {
    db: db ?? {},
    authSession: null,
    user: null,
  };
}

export function createTrpcCaller(ctx: MinimalContext) {
  return appRouter.createCaller(ctx as never);
}
