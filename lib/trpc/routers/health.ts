import { router, publicProcedure } from 'lib/trpc/trpc';

export const healthRouter = router({
  hello: publicProcedure.query(() => ({ message: 'Hello from tRPC!' })),
  ping: publicProcedure.query(() => ({ message: 'pong' })),
});
