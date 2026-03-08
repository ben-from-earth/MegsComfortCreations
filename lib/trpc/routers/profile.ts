import { router, adminProcedure } from 'lib/trpc/trpc';

export const profileRouter = router({
  get: adminProcedure.query(async () => {
    const user = {
      id: 123,
      firstName: 'Ben',
      lastName: 'Knox',
      email: 'example@email.com',
    };
    return user;
  }),
});
