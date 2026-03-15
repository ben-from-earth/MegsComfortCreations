import { Db } from '@/db/client';
import { books, otherMedia } from '@/db/schema';
import { eq } from 'drizzle-orm';
import {
  getImageMigrationStatus,
  replaceBookImageRecords,
  replaceOtherMediaImageRecords,
  resolveAndPersistImageList,
} from './media-image-records';
import { isExternalImageUrl } from './image-path-utils';

export type LegacyImageMigrationSummary = {
  dryRun: boolean;
  processedItems: number;
  migratedExternalUrls: number;
  skippedLocalPaths: number;
  failedDownloads: number;
  deletedRows: number;
  failures: Array<{
    mediaId: string;
    mediaType: string;
    sourceUrl: string;
    message: string;
  }>;
};

type LegacyImageMigrationInput = {
  db: Db;
  dryRun?: boolean;
};

export async function migrateLegacyImageUrlsToLocalFiles(
  input: LegacyImageMigrationInput,
): Promise<LegacyImageMigrationSummary> {
  const dryRun = input.dryRun ?? false;
  const db = input.db;

  const summary: LegacyImageMigrationSummary = {
    dryRun,
    processedItems: 0,
    migratedExternalUrls: 0,
    skippedLocalPaths: 0,
    failedDownloads: 0,
    deletedRows: 0,
    failures: [],
  };

  const [bookRows, otherMediaRows] = await Promise.all([
    db.select().from(books),
    db.select().from(otherMedia),
  ]);

  for (const bookRow of bookRows) {
    summary.processedItems += 1;
    const sourceImageUrls = bookRow.imageUrls;
    const externalSourceUrls = sourceImageUrls.filter((imageUrl) =>
      isExternalImageUrl(imageUrl),
    );
    const localOrManagedSourceCount =
      sourceImageUrls.length - externalSourceUrls.length;
    const resolved = await resolveAndPersistImageList(
      { mediaType: 'book', mediaId: bookRow.id },
      sourceImageUrls,
    );

    summary.failedDownloads += resolved.failures.length;
    summary.migratedExternalUrls +=
      externalSourceUrls.length - resolved.failures.length;
    summary.skippedLocalPaths += localOrManagedSourceCount;
    for (const failure of resolved.failures) {
      summary.failures.push({
        mediaId: bookRow.id,
        mediaType: 'book',
        sourceUrl: failure.sourceUrl,
        message: failure.message,
      });
    }

    const hasFailures = resolved.failures.length > 0;
    if (!dryRun && hasFailures) {
      await db.delete(books).where(eq(books.id, bookRow.id));
      summary.deletedRows += 1;
      continue;
    }

    if (!dryRun) {
      await replaceBookImageRecords(db, bookRow.id, resolved.images);
    }
  }

  for (const mediaRow of otherMediaRows) {
    summary.processedItems += 1;
    const sourceImageUrls = mediaRow.imageUrls;
    const externalSourceUrls = sourceImageUrls.filter((imageUrl) =>
      isExternalImageUrl(imageUrl),
    );
    const localOrManagedSourceCount =
      sourceImageUrls.length - externalSourceUrls.length;
    const resolved = await resolveAndPersistImageList(
      { mediaType: mediaRow.mediaType, mediaId: mediaRow.id },
      sourceImageUrls,
    );

    summary.failedDownloads += resolved.failures.length;
    summary.migratedExternalUrls +=
      externalSourceUrls.length - resolved.failures.length;
    summary.skippedLocalPaths += localOrManagedSourceCount;
    for (const failure of resolved.failures) {
      summary.failures.push({
        mediaId: mediaRow.id,
        mediaType: mediaRow.mediaType,
        sourceUrl: failure.sourceUrl,
        message: failure.message,
      });
    }

    const hasFailures = resolved.failures.length > 0;
    if (!dryRun && hasFailures) {
      await db.delete(otherMedia).where(eq(otherMedia.id, mediaRow.id));
      summary.deletedRows += 1;
      continue;
    }

    if (!dryRun) {
      await replaceOtherMediaImageRecords(db, mediaRow.id, resolved.images);
    }
  }

  return summary;
}

export async function runOneTimeLegacyImageMigration(db: Db) {
  const statusBefore = await getImageMigrationStatus(db);
  if (statusBefore.isCompleted) {
    return {
      alreadyCompleted: true,
      statusBefore,
      statusAfter: statusBefore,
      summary: null,
    };
  }

  const summary = await migrateLegacyImageUrlsToLocalFiles({
    db,
    dryRun: false,
  });
  const statusAfter = await getImageMigrationStatus(db);
  return {
    alreadyCompleted: false,
    statusBefore,
    statusAfter,
    summary,
  };
}
