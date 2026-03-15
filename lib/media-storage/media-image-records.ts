import { Db } from '@/db/client';
import { and, asc, eq, inArray, isNull } from 'drizzle-orm';
import { books, mediaImageItems, otherMedia } from '@/db/schema';
import { MediaType } from 'lib/constants/mediaTypes';
import { isExternalImageUrl, normalizeImagePath } from './image-path-utils';
import {
  PersistedImageFile,
  persistExternalImageToS3,
} from './local-image-storage';

type ParentMediaReference =
  | { mediaType: 'book'; mediaId: string }
  | { mediaType: Exclude<MediaType, 'book'>; mediaId: string };

export type ImageResolutionResult = {
  images: PersistedImageFile[];
  failures: Array<{ sourceUrl: string; message: string }>;
};

export async function resolveAndPersistImageList(
  mediaReference: ParentMediaReference,
  sourceImageUrls: string[],
): Promise<ImageResolutionResult> {
  const images: PersistedImageFile[] = [];
  const failures: Array<{ sourceUrl: string; message: string }> = [];

  for (let index = 0; index < sourceImageUrls.length; index += 1) {
    const sourceUrl = normalizeImagePath(sourceImageUrls[index] ?? '');
    if (!sourceUrl) {
      continue;
    }

    try {
      const persistedImage = await persistExternalImageToS3({
        sourceUrl,
        mediaType: mediaReference.mediaType,
        mediaId: mediaReference.mediaId,
        sortOrder: index,
      });
      images.push(persistedImage);
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'Unknown image persistence error';
      failures.push({ sourceUrl, message });
    }
  }

  return { images, failures };
}

export async function replaceBookImageRecords(
  db: Db,
  bookId: string,
  imageFiles: PersistedImageFile[],
) {
  await db.delete(mediaImageItems).where(eq(mediaImageItems.bookId, bookId));

  if (imageFiles.length === 0) {
    return;
  }

  await db.insert(mediaImageItems).values(
    imageFiles.map((imageFile, sortOrder) => ({
      bookId,
      path: imageFile.publicPath,
      sourceUrl: imageFile.sourceUrl,
      mimeType: imageFile.mimeType,
      sizeBytes: imageFile.sizeBytes,
      sortOrder,
    })),
  );
}

export async function replaceOtherMediaImageRecords(
  db: Db,
  otherMediaId: string,
  imageFiles: PersistedImageFile[],
) {
  await db
    .delete(mediaImageItems)
    .where(eq(mediaImageItems.otherMediaId, otherMediaId));

  if (imageFiles.length === 0) {
    return;
  }

  await db.insert(mediaImageItems).values(
    imageFiles.map((imageFile, sortOrder) => ({
      otherMediaId,
      path: imageFile.publicPath,
      sourceUrl: imageFile.sourceUrl,
      mimeType: imageFile.mimeType,
      sizeBytes: imageFile.sizeBytes,
      sortOrder,
    })),
  );
}

export async function loadBookImageUrlsById(db: Db, bookIds: string[]) {
  if (bookIds.length === 0) {
    return new Map<string, string[]>();
  }

  const rows = await db
    .select({
      bookId: mediaImageItems.bookId,
      path: mediaImageItems.path,
      sortOrder: mediaImageItems.sortOrder,
    })
    .from(mediaImageItems)
    .where(
      and(
        inArray(mediaImageItems.bookId, bookIds),
        isNull(mediaImageItems.otherMediaId),
      ),
    )
    .orderBy(asc(mediaImageItems.sortOrder));

  const imageUrlsByBookId = new Map<string, string[]>();
  for (const row of rows) {
    const bookId = row.bookId;
    if (!bookId) {
      continue;
    }
    const currentList = imageUrlsByBookId.get(bookId) ?? [];
    currentList.push(row.path);
    imageUrlsByBookId.set(bookId, currentList);
  }
  return imageUrlsByBookId;
}

export async function loadOtherMediaImageUrlsById(db: Db, otherMediaIds: string[]) {
  if (otherMediaIds.length === 0) {
    return new Map<string, string[]>();
  }

  const rows = await db
    .select({
      otherMediaId: mediaImageItems.otherMediaId,
      path: mediaImageItems.path,
      sortOrder: mediaImageItems.sortOrder,
    })
    .from(mediaImageItems)
    .where(
      and(
        inArray(mediaImageItems.otherMediaId, otherMediaIds),
        isNull(mediaImageItems.bookId),
      ),
    )
    .orderBy(asc(mediaImageItems.sortOrder));

  const imageUrlsByOtherMediaId = new Map<string, string[]>();
  for (const row of rows) {
    const otherMediaId = row.otherMediaId;
    if (!otherMediaId) {
      continue;
    }
    const currentList = imageUrlsByOtherMediaId.get(otherMediaId) ?? [];
    currentList.push(row.path);
    imageUrlsByOtherMediaId.set(otherMediaId, currentList);
  }
  return imageUrlsByOtherMediaId;
}

export type MigrationStatus = {
  totalItems: number;
  externalUrlCount: number;
  missingReferenceCount: number;
  isCompleted: boolean;
};

export async function getImageMigrationStatus(db: Db): Promise<MigrationStatus> {
  const [bookRowsWithIds, otherMediaRowsWithIds, imageReferenceRows] = await Promise.all([
    db.select({ id: books.id, imageUrls: books.imageUrls }).from(books),
    db.select({ id: otherMedia.id, imageUrls: otherMedia.imageUrls }).from(otherMedia),
    db
      .select({
        bookId: mediaImageItems.bookId,
        otherMediaId: mediaImageItems.otherMediaId,
      })
      .from(mediaImageItems),
  ]);

  let externalUrlCount = 0;
  for (const row of [...bookRowsWithIds, ...otherMediaRowsWithIds]) {
    for (const imageUrl of row.imageUrls) {
      if (isExternalImageUrl(imageUrl)) {
        externalUrlCount += 1;
      }
    }
  }

  const referencedBookIds = new Set(
    imageReferenceRows
      .map((row) => row.bookId)
      .filter((bookId): bookId is string => typeof bookId === 'string'),
  );
  const referencedOtherMediaIds = new Set(
    imageReferenceRows
      .map((row) => row.otherMediaId)
      .filter((otherMediaId): otherMediaId is string => typeof otherMediaId === 'string'),
  );

  let missingReferenceCount = 0;
  for (const row of bookRowsWithIds) {
    if (row.imageUrls.length > 0 && !referencedBookIds.has(row.id)) {
      missingReferenceCount += 1;
    }
  }
  for (const row of otherMediaRowsWithIds) {
    if (row.imageUrls.length > 0 && !referencedOtherMediaIds.has(row.id)) {
      missingReferenceCount += 1;
    }
  }

  const totalItems = bookRowsWithIds.length + otherMediaRowsWithIds.length;

  return {
    totalItems,
    externalUrlCount,
    missingReferenceCount,
    isCompleted: externalUrlCount === 0 && missingReferenceCount === 0,
  };
}
