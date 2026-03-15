import { createAdminContext, createTrpcCaller } from '../helpers/trpcTestContext';

describe('profile router', () => {
  test('get returns expected static profile contract', async () => {
    const caller = createTrpcCaller(createAdminContext({}));
    const profile = await caller.profile.get();

    expect(profile).toEqual({
      id: 123,
      firstName: 'Ben',
      lastName: 'Knox',
      email: 'example@email.com',
    });
  });
});
