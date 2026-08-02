import { createAnonContext, createTrpcCaller, createUserContext } from '../helpers/trpcTestContext';

describe('tRPC auth guards', () => {
  test('admin endpoint rejects anonymous callers', async () => {
    const caller = createTrpcCaller(createAnonContext({}));
    await expect(caller.genres.getAll()).rejects.toMatchObject({
      code: 'UNAUTHORIZED',
      message: 'Login required',
    });
  });

  test('admin endpoint rejects non-admin users', async () => {
    const caller = createTrpcCaller(createUserContext({}));
    await expect(caller.genres.getAll()).rejects.toMatchObject({
      code: 'FORBIDDEN',
      message: 'Admin role required',
    });
  });
});
