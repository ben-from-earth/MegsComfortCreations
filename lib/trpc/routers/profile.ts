import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import {
  getImageMigrationStatus,
} from 'lib/media-storage/media-image-records';
import {
  migrateLegacyImageUrlsToLocalFiles,
  runOneTimeLegacyImageMigration,
} from 'lib/media-storage/migrate-legacy-image-urls';

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
  getImageMigrationStatus: adminProcedure.query(async ({ ctx }) => {
    const db = ctx.db ?? defaultDb;
    return getImageMigrationStatus(db);
  }),
  migrateImageFiles: adminProcedure
    .input(z.object({ dryRun: z.boolean().default(false) }).optional())
    .mutation(async ({ ctx, input }) => {
      const db = ctx.db ?? defaultDb;
      const dryRun = input?.dryRun ?? false;

      if (dryRun) {
        const statusBefore = await getImageMigrationStatus(db);
        const summary = await migrateLegacyImageUrlsToLocalFiles({ db, dryRun: true });
        return {
          dryRun: true,
          alreadyCompleted: statusBefore.isCompleted,
          statusBefore,
          statusAfter: statusBefore,
          summary,
        };
      }

      const migrationResult = await runOneTimeLegacyImageMigration(db);
      return {
        dryRun: false,
        ...migrationResult,
      };
    }),
});
