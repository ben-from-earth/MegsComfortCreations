import { router } from 'lib/trpc/trpc';
import { databaseRouter } from './database';
import { genresRouter } from './genres';
import { onlineRouter } from './online';
import { pngRouter } from './png';
import { profileRouter } from './profile';
import { healthRouter } from './health';

export const appRouter = router({
  database: databaseRouter,
  genres: genresRouter,
  online: onlineRouter,
  png: pngRouter,
  profile: profileRouter,
  health: healthRouter,
});

export type AppRouter = typeof appRouter;
