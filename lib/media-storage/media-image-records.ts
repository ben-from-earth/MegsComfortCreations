import { Db } from '@/db/client';
import { asc, eq, inArray } from 'drizzle-orm';
import { mediaImageItems } from '@/db/schema';
import { MediaType } from 'lib/constants/media-types';
import { normalizeImagePath } from './image-path-utils';
import { MediaImageItem } from 'lib/interfaces/global-interfaces';
import {
  PersistedImageFile,
  persistExternalImageToS3,
} from './local-image-storage';

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
    };

export type PersistedMediaImageRecord = PersistedImageFile & {
  isDefault: boolean;
  spineColor: string;
  codeColor: string | null;
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
        return {
          image: {
            ...persistedImage,
            isDefault: index === resolvedDefaultIndex,
            spineColor,
            codeColor,
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

export async function replaceBookImageRecords(
  db: Db,
  bookId: string,
  imageFiles: PersistedMediaImageRecord[],
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
      isDefault: imageFile.isDefault,
      spineColor: imageFile.spineColor,
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

export async function replaceMediaImageRecords(
  db: Db,
  mediaReference: ParentMediaReference,
  imageFiles: PersistedMediaImageRecord[],
) {
  const mediaId = mediaReference.mediaId;
  switch (mediaReference.mediaType) {
    case 'book':
      await db.delete(mediaImageItems).where(eq(mediaImageItems.bookId, mediaId));
      break;
    case 'movie':
      await db
        .delete(mediaImageItems)
        .where(eq(mediaImageItems.movieId, mediaId));
      break;
    case 'videoGame':
      await db
        .delete(mediaImageItems)
        .where(eq(mediaImageItems.videoGameId, mediaId));
      break;
    case 'album':
      await db
        .delete(mediaImageItems)
        .where(eq(mediaImageItems.albumId, mediaId));
      break;
  }

  if (imageFiles.length === 0) {
    return;
  }

  const codeColorFor = (imageFile: PersistedMediaImageRecord) =>
    mediaReference.mediaType === 'album' ? imageFile.codeColor : null;

  switch (mediaReference.mediaType) {
    case 'book':
      await db.insert(mediaImageItems).values(
        imageFiles.map((imageFile, sortOrder) => ({
          bookId: mediaId,
          ...imageRecordValues(imageFile, sortOrder, codeColorFor(imageFile)),
        })),
      );
      return;
    case 'movie':
      await db.insert(mediaImageItems).values(
        imageFiles.map((imageFile, sortOrder) => ({
          movieId: mediaId,
          ...imageRecordValues(imageFile, sortOrder, codeColorFor(imageFile)),
        })),
      );
      return;
    case 'videoGame':
      await db.insert(mediaImageItems).values(
        imageFiles.map((imageFile, sortOrder) => ({
          videoGameId: mediaId,
          ...imageRecordValues(imageFile, sortOrder, codeColorFor(imageFile)),
        })),
      );
      return;
    case 'album':
      await db.insert(mediaImageItems).values(
        imageFiles.map((imageFile, sortOrder) => ({
          albumId: mediaId,
          ...imageRecordValues(imageFile, sortOrder, codeColorFor(imageFile)),
        })),
      );
      return;
  }
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
