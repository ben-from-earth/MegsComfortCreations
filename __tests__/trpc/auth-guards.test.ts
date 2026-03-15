import { createAnonContext, createTrpcCaller, createUserContext } from '../helpers/trpcTestContext';

describe('tRPC auth guards', () => {
  test('public health endpoint is available without auth', async () => {
    const caller = createTrpcCaller(createAnonContext({}));
    const response = await caller.health.ping();
    expect(response).toEqual({ message: 'pong' });
  });

  test('admin endpoint rejects anonymous callers', async () => {
    const caller = createTrpcCaller(createAnonContext({}));
    await expect(caller.profile.get()).rejects.toMatchObject({
      code: 'UNAUTHORIZED',
      message: 'Login required',
    });
  });

  test('admin endpoint rejects non-admin users', async () => {
    const caller = createTrpcCaller(createUserContext({}));
    await expect(caller.profile.get()).rejects.toMatchObject({
      code: 'FORBIDDEN',
      message: 'Admin role required',
    });
  });
});
