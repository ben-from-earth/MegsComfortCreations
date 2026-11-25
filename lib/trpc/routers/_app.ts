import { router } from 'lib/trpc/trpc';
import { genresRouter } from './genres/_';
import { pngRouter } from './png';
import { databaseRouter } from './database/_';
import { collectRouter } from './collect/_';
import { ordersRouter } from './orders/_';

export const appRouter = router({
  database: databaseRouter,
  collect: collectRouter,
  genres: genresRouter,
  png: pngRouter,
  orders: ordersRouter,
});

export type AppRouter = typeof appRouter;
