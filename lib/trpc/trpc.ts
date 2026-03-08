import { initTRPC } from '@trpc/server';
import { TRPCError } from '@trpc/server';
import superjson from 'superjson';
import type { Context } from './context';

const t = initTRPC.context<Context>().create({
  transformer: superjson,
});

export const router = t.router;
export const publicProcedure = t.procedure;

export const protectedProcedure = t.procedure.use(({ ctx, next }) => {
  if (!ctx.authSession || !ctx.user) {
    throw new TRPCError({ code: 'UNAUTHORIZED', message: 'Login required' });
  }

  return next({
    ctx: {
      ...ctx,
      authSession: ctx.authSession,
      user: ctx.user,
    },
  });
});

export const adminProcedure = protectedProcedure.use(({ ctx, next }) => {
  if (ctx.user.role !== 'admin') {
    throw new TRPCError({
      code: 'FORBIDDEN',
      message: 'Admin role required',
    });
  }

  return next();
});
