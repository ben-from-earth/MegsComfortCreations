type TestUser = {
  id: string;
  role: 'admin' | 'user';
};

/**
 * Auth identity for slim procedure-guard tests.
 * Live Context also carries full Better Auth session / Drizzle client shapes.
 */
export type TrpcTestContext = {
  db: object;
  authSession: { user: TestUser } | null;
  user: TestUser | null;
};

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
