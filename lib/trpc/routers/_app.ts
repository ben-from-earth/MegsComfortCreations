import { router } from 'lib/trpc/trpc';
import { genresRouter } from './genres/_';
import { onlineRouter } from './online';
import { pngRouter } from './png';
import { healthRouter } from './health';
import { databaseRouter } from './database/_';
import { collectRouter } from './collect/_';

export const appRouter = router({
  database: databaseRouter,
  collect: collectRouter,
  genres: genresRouter,
  online: onlineRouter,
  png: pngRouter,
  health: healthRouter,
});

export type AppRouter = typeof appRouter;
