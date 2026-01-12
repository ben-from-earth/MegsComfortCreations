import { router, publicProcedure } from 'lib/trpc/trpc';

export const profileRouter = router({
  get: publicProcedure.query(async () => {
    const user = {
      id: 123,
      firstName: 'Ben',
      lastName: 'Knox',
      email: 'example@email.com',
    };
    return user;
  }),
});
