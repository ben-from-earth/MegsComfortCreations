import { adminProcedure, router } from 'lib/trpc/trpc';
import type { Context } from 'lib/trpc/context';
import {
  createAnonContext,
  createUserContext,
  type TrpcTestContext,
} from '../helpers/trpc-test-context';

const authGuardRouter = router({
  adminPing: adminProcedure.query(() => true),
});

function createCaller(ctx: TrpcTestContext) {
  return authGuardRouter.createCaller(ctx as Context);
}

describe('tRPC auth guards', () => {
  test('admin procedure rejects anonymous callers', async () => {
    await expect(createCaller(createAnonContext()).adminPing()).rejects.toMatchObject({
      code: 'UNAUTHORIZED',
      message: 'Login required',
    });
  });

  test('admin procedure rejects non-admin users', async () => {
    await expect(createCaller(createUserContext()).adminPing()).rejects.toMatchObject({
      code: 'FORBIDDEN',
      message: 'Admin role required',
    });
  });
});
