import { Db, DbExecutor } from '@/db/client';
import { and, asc, eq, inArray, ne } from 'drizzle-orm';
import { mediaImageItems } from '@/db/schema';
import { MediaType } from 'lib/constants/media-types';
import { normalizeImagePath } from './image-path-utils';
import { MediaImageItem } from 'lib/interfaces/global-interfaces';
import {
  PersistedImageFile,
  persistExternalImageToS3,
} from './local-image-storage';
import { planCoverRecordChanges } from './plan-cover-record-changes';

type ParentMediaReference =
  | { mediaType: 'book'; mediaId: string }
  | { mediaType: Exclude<MediaType, 'book'>; mediaId: string };

export type ImageResolutionResult = {
  images: PersistedMediaImageRecord[];
  failures: Array<{ sourceUrl: string; message: string }>;
};

export type SourceImageItem =
  | string
  | {
      url: string;
      isDefault?: boolean;
      spineColor?: string;
      codeColor?: string | null;
      id?: string;
    };

export type PersistedMediaImageRecord = PersistedImageFile & {
  isDefault: boolean;
  spineColor: string;
  codeColor: string | null;
  id?: string;
};

export type SavedCoverRecord = PersistedMediaImageRecord & {
  id: string;
};

function normalizeSourceImage(
  rawSourceImage: unknown,
  fallbackSpineColor: string,
): SourceImageItem | null {
  if (typeof rawSourceImage === 'string') {
    return rawSourceImage;
  }
  if (typeof rawSourceImage === 'object' && rawSourceImage !== null) {
    const imageRecord = rawSourceImage as {
      url?: unknown;
      src?: unknown;
      isDefault?: boolean;
      spineColor?: string;
      codeColor?: unknown;
      id?: unknown;
    };
    const url =
      typeof imageRecord.url === 'string'
        ? imageRecord.url
        : typeof imageRecord.src === 'string'
          ? imageRecord.src
          : null;
    if (url !== null) {
      return {
        url,
        isDefault: imageRecord.isDefault ?? false,
        spineColor: imageRecord.spineColor ?? fallbackSpineColor,
        codeColor:
          typeof imageRecord.codeColor === 'string' ? imageRecord.codeColor : null,
        ...(typeof imageRecord.id === 'string' && imageRecord.id.length > 0
          ? { id: imageRecord.id }
          : {}),
      };
    }
  }
  return null;
}

export function coversToPersistForNewItem<
  T extends { selected: boolean; isDefault: boolean; spineColor: string },
>(images: T[]): T[] {
  const keptImages = images.filter((image) => image.selected);
  const covers =
    keptImages.length > 0
      ? keptImages
      : images.length > 0
        ? [{ ...images[0], selected: true, isDefault: true }]
        : [];
  const explicitDefaultIndex = covers.findIndex((image) => image.isDefault);
  const defaultIndex = explicitDefaultIndex >= 0 ? explicitDefaultIndex : 0;

  return covers.map((image, index) => ({
    ...image,
    isDefault: index === defaultIndex,
  }));
}

export async function resolveAndPersistImageList(
  mediaReference: ParentMediaReference,
  sourceImages: unknown[],
  options?: { defaultSpineColor?: string; defaultImageIndex?: number },
): Promise<ImageResolutionResult> {
  const images: PersistedMediaImageRecord[] = [];
  const failures: Array<{ sourceUrl: string; message: string }> = [];
  const fallbackSpineColor = options?.defaultSpineColor ?? '#ffffff';
  const defaultImageIndex = options?.defaultImageIndex ?? 0;

  const normalizedSourceImages: SourceImageItem[] = [];
  for (const sourceImage of sourceImages) {
    const normalizedSourceImage = normalizeSourceImage(
      sourceImage,
      fallbackSpineColor,
    );
    if (normalizedSourceImage === null) {
      if (
        sourceImage !== null &&
        sourceImage !== undefined &&
        sourceImage !== ''
      ) {
        failures.push({
          sourceUrl: '',
          message: 'Invalid image payload. Expected an image URL string.',
        });
      }
      continue;
    }
    normalizedSourceImages.push(normalizedSourceImage);
  }
  const explicitDefaultIndex = normalizedSourceImages.findIndex(
    (sourceImage) =>
      typeof sourceImage === 'object' &&
      sourceImage !== null &&
      sourceImage.isDefault === true,
  );
  const resolvedDefaultIndex =
    explicitDefaultIndex >= 0 ? explicitDefaultIndex : defaultImageIndex;

  const persistResults = await Promise.all(
    normalizedSourceImages.map(async (
      rawSourceImage,
      index,
    ): Promise<
      | { image: PersistedMediaImageRecord }
      | { failure: { sourceUrl: string; message: string } }
      | null
    > => {
      const sourceUrl = normalizeImagePath(
        typeof rawSourceImage === 'string'
          ? rawSourceImage
          : rawSourceImage.url,
      );
      if (!sourceUrl) {
        return rawSourceImage === ''
          ? null
          : {
              failure: {
                sourceUrl: '',
                message: 'Invalid image payload. Expected an image URL string.',
              },
            };
      }

      try {
        const persistedImage = await persistExternalImageToS3({
          sourceUrl,
          mediaType: mediaReference.mediaType,
          mediaId: mediaReference.mediaId,
          sortOrder: index,
        });
        const spineColor =
          typeof rawSourceImage === 'object' && rawSourceImage.spineColor
            ? rawSourceImage.spineColor
            : fallbackSpineColor;
        const codeColor =
          typeof rawSourceImage === 'object' ? (rawSourceImage.codeColor ?? null) : null;
        const id =
          typeof rawSourceImage === 'object' ? rawSourceImage.id : undefined;
        return {
          image: {
            ...persistedImage,
            isDefault: index === resolvedDefaultIndex,
            spineColor,
            codeColor,
            ...(id ? { id } : {}),
          },
        };
      } catch (error) {
        const message =
          error instanceof Error
            ? error.message
            : 'Unknown image persistence error';
        return { failure: { sourceUrl, message } };
      }
    }),
  );

  for (const persistResult of persistResults) {
    if (persistResult === null) {
      continue;
    }
    if ('image' in persistResult) {
      images.push(persistResult.image);
    } else {
      failures.push(persistResult.failure);
    }
  }

  return { images, failures };
}

export function listCoversToAppend(
  storedImages: readonly { url: string }[],
  uploadedImages: readonly { url: string; spineColor: string }[],
): { url: string; isDefault: boolean; spineColor: string }[] {
  const storedUrls = new Set(storedImages.map((image) => image.url));
  const coversToAppend: {
    url: string;
    isDefault: boolean;
    spineColor: string;
  }[] = [];

  for (const image of uploadedImages) {
    const url = image.url.trim();
    if (!url || storedUrls.has(url)) {
      continue;
    }
    storedUrls.add(url);
    coversToAppend.push({
      url,
      isDefault: storedImages.length === 0 && coversToAppend.length === 0,
      spineColor: image.spineColor,
    });
  }

  return coversToAppend;
}

export async function appendBookCoverRecords(
  db: Db,
  bookId: string,
  covers: readonly { url: string; isDefault: boolean; spineColor: string }[],
) {
  if (covers.length === 0) {
    return;
  }

  const existingRows = await db
    .select({ sortOrder: mediaImageItems.sortOrder })
    .from(mediaImageItems)
    .where(eq(mediaImageItems.bookId, bookId));

  let nextSortOrder = 0;
  for (const row of existingRows) {
    if (row.sortOrder + 1 > nextSortOrder) {
      nextSortOrder = row.sortOrder + 1;
    }
  }

  await db.insert(mediaImageItems).values(
    covers.map((cover, index) => ({
      bookId,
      path: cover.url,
      sourceUrl: cover.url,
      sortOrder: nextSortOrder + index,
      isDefault: cover.isDefault,
      spineColor: cover.spineColor,
    })),
  );
}

function imageRecordValues(
  imageFile: PersistedMediaImageRecord,
  sortOrder: number,
  codeColor: string | null,
) {
  return {
    path: imageFile.publicPath,
    sourceUrl: imageFile.sourceUrl,
    mimeType: imageFile.mimeType,
    sizeBytes: imageFile.sizeBytes,
    sortOrder,
    isDefault: imageFile.isDefault,
    spineColor: imageFile.spineColor,
    codeColor,
  };
}

function parentMatch(mediaType: MediaType, mediaId: string) {
  switch (mediaType) {
    case 'book':
      return eq(mediaImageItems.bookId, mediaId);
    case 'movie':
      return eq(mediaImageItems.movieId, mediaId);
    case 'videoGame':
      return eq(mediaImageItems.videoGameId, mediaId);
    case 'album':
      return eq(mediaImageItems.albumId, mediaId);
  }
}

function storedCodeColor(
  mediaType: MediaType,
  imageFile: PersistedMediaImageRecord,
) {
  return mediaType === 'album' ? imageFile.codeColor : null;
}

async function insertCover(
  db: DbExecutor,
  mediaType: MediaType,
  mediaId: string,
  imageFile: PersistedMediaImageRecord,
  sortOrder: number,
) {
  const values = imageRecordValues(
    imageFile,
    sortOrder,
    storedCodeColor(mediaType, imageFile),
  );
  switch (mediaType) {
    case 'book':
      return db
        .insert(mediaImageItems)
        .values({ bookId: mediaId, ...values })
        .returning({ id: mediaImageItems.id });
    case 'movie':
      return db
        .insert(mediaImageItems)
        .values({ movieId: mediaId, ...values })
        .returning({ id: mediaImageItems.id });
    case 'videoGame':
      return db
        .insert(mediaImageItems)
        .values({ videoGameId: mediaId, ...values })
        .returning({ id: mediaImageItems.id });
    case 'album':
      return db
        .insert(mediaImageItems)
        .values({ albumId: mediaId, ...values })
        .returning({ id: mediaImageItems.id });
  }
}

async function replaceCoverRecords(
  db: DbExecutor,
  mediaType: MediaType,
  mediaId: string,
  imageFiles: PersistedMediaImageRecord[],
): Promise<SavedCoverRecord[]> {
  const parentWhere = parentMatch(mediaType, mediaId);
  const existingRows = await db
    .select({ id: mediaImageItems.id })
    .from(mediaImageItems)
    .where(parentWhere);
  const plan = planCoverRecordChanges(
    existingRows.map((row) => row.id),
    imageFiles,
  );

  if (imageFiles.length === 0) {
    if (plan.deletes.length > 0) {
      await db
        .delete(mediaImageItems)
        .where(and(parentWhere, inArray(mediaImageItems.id, plan.deletes)));
    }
    return [];
  }

  const defaultIndex = imageFiles.findIndex((image) => image.isDefault);
  const keptDefault = plan.updates.find(
    (update) => update.incomingIndex === defaultIndex,
  );
  if (existingRows.length > 0) {
    await db
      .update(mediaImageItems)
      .set({ isDefault: false })
      .where(
        keptDefault
          ? and(parentWhere, ne(mediaImageItems.id, keptDefault.id))
          : parentWhere,
      );
  }

  const saved: SavedCoverRecord[] = new Array(imageFiles.length);
  for (const update of plan.updates) {
    const imageFile = imageFiles[update.incomingIndex];
    const codeColor = storedCodeColor(mediaType, imageFile);
    await db
      .update(mediaImageItems)
      .set(imageRecordValues(imageFile, update.incomingIndex, codeColor))
      .where(eq(mediaImageItems.id, update.id));
    saved[update.incomingIndex] = { ...imageFile, id: update.id, codeColor };
  }

  for (const insert of plan.inserts) {
    const imageFile = imageFiles[insert.incomingIndex];
    const [row] = await insertCover(
      db,
      mediaType,
      mediaId,
      imageFile,
      insert.incomingIndex,
    );
    if (!row) {
      throw new Error('Cover insert failed');
    }
    saved[insert.incomingIndex] = {
      ...imageFile,
      id: row.id,
      codeColor: storedCodeColor(mediaType, imageFile),
    };
  }

  if (plan.deletes.length > 0) {
    await db
      .delete(mediaImageItems)
      .where(and(parentWhere, inArray(mediaImageItems.id, plan.deletes)));
  }

  return saved;
}

export async function replaceBookImageRecords(
  db: DbExecutor,
  bookId: string,
  imageFiles: PersistedMediaImageRecord[],
) {
  return replaceCoverRecords(db, 'book', bookId, imageFiles);
}

export async function replaceMediaImageRecords(
  db: DbExecutor,
  mediaReference: ParentMediaReference,
  imageFiles: PersistedMediaImageRecord[],
) {
  return replaceCoverRecords(
    db,
    mediaReference.mediaType,
    mediaReference.mediaId,
    imageFiles,
  );
}

export async function loadBookImagesById(db: Db, bookIds: string[]) {
  if (bookIds.length === 0) {
    return new Map<string, MediaImageItem[]>();
  }

  return loadImagesByParentColumn(db, mediaImageItems.bookId, bookIds);
}

async function loadImagesByParentColumn(
  db: Db,
  parentIdColumn:
    | typeof mediaImageItems.bookId
    | typeof mediaImageItems.movieId
    | typeof mediaImageItems.videoGameId
    | typeof mediaImageItems.albumId,
  parentIds: string[],
) {
  if (parentIds.length === 0) {
    return new Map<string, MediaImageItem[]>();
  }

  const rows = await db
    .select({
      id: mediaImageItems.id,
      parentId: parentIdColumn,
      path: mediaImageItems.path,
      sortOrder: mediaImageItems.sortOrder,
      isDefault: mediaImageItems.isDefault,
      spineColor: mediaImageItems.spineColor,
      codeColor: mediaImageItems.codeColor,
    })
    .from(mediaImageItems)
    .where(inArray(parentIdColumn, parentIds))
    .orderBy(asc(mediaImageItems.sortOrder));

  const imagesByParentId = new Map<string, MediaImageItem[]>();
  for (const row of rows) {
    const parentId = row.parentId;
    if (!parentId) {
      continue;
    }
    const currentList = imagesByParentId.get(parentId) ?? [];
    currentList.push({
      id: row.id,
      url: row.path,
      isDefault: row.isDefault,
      spineColor: row.spineColor,
      codeColor: row.codeColor,
      selected: false,
    });
    imagesByParentId.set(parentId, currentList);
  }
  return imagesByParentId;
}

export function loadMovieImagesById(db: Db, movieIds: string[]) {
  return loadImagesByParentColumn(db, mediaImageItems.movieId, movieIds);
}

export function loadVideoGameImagesById(db: Db, videoGameIds: string[]) {
  return loadImagesByParentColumn(db, mediaImageItems.videoGameId, videoGameIds);
}

export function loadAlbumImagesById(db: Db, albumIds: string[]) {
  return loadImagesByParentColumn(db, mediaImageItems.albumId, albumIds);
}
