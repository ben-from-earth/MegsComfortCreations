import { createAnonContext, createTrpcCaller } from '../helpers/trpcTestContext';

describe('health router', () => {
  test('hello returns expected payload', async () => {
    const caller = createTrpcCaller(createAnonContext({}));
    await expect(caller.health.hello()).resolves.toEqual({
      message: 'Hello from tRPC!',
    });
  });

  test('ping returns pong', async () => {
    const caller = createTrpcCaller(createAnonContext({}));
    await expect(caller.health.ping()).resolves.toEqual({ message: 'pong' });
  });
});
