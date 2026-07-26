import { Db } from '@/db/client';
import { and, asc, eq, inArray, isNull } from 'drizzle-orm';
import { mediaImageItems } from '@/db/schema';
import { MediaType } from 'lib/constants/mediaTypes';
import { normalizeImagePath } from './image-path-utils';
import { MediaImageItem } from 'lib/interfaces/globalInterfaces';
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
    };

export type PersistedMediaImageRecord = PersistedImageFile & {
  isDefault: boolean;
  spineColor: string;
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
      };
    }
  }
  return null;
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
      if (!(sourceImage == null || sourceImage === '')) {
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
  const resolvedDefaultIndex = explicitDefaultIndex >= 0 ? explicitDefaultIndex : defaultImageIndex;

  for (let index = 0; index < normalizedSourceImages.length; index += 1) {
    const rawSourceImage = normalizedSourceImages[index];
    const sourceUrl = normalizeImagePath(
      typeof rawSourceImage === 'string' ? rawSourceImage : rawSourceImage.url,
    );
    if (!sourceUrl) {
      if (!(rawSourceImage == null || rawSourceImage === '')) {
        failures.push({
          sourceUrl: '',
          message: 'Invalid image payload. Expected an image URL string.',
        });
      }
      continue;
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
      images.push({
        ...persistedImage,
        isDefault: index === resolvedDefaultIndex,
        spineColor,
      });
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

export async function replaceOtherMediaImageRecords(
  db: Db,
  otherMediaId: string,
  imageFiles: PersistedMediaImageRecord[],
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
      isDefault: imageFile.isDefault,
      spineColor: imageFile.spineColor,
    })),
  );
}

export async function loadBookImagesById(db: Db, bookIds: string[]) {
  if (bookIds.length === 0) {
    return new Map<string, MediaImageItem[]>();
  }

  const rows = await db
    .select({
      bookId: mediaImageItems.bookId,
      path: mediaImageItems.path,
      sortOrder: mediaImageItems.sortOrder,
      isDefault: mediaImageItems.isDefault,
      spineColor: mediaImageItems.spineColor,
    })
    .from(mediaImageItems)
    .where(
      and(
        inArray(mediaImageItems.bookId, bookIds),
        isNull(mediaImageItems.otherMediaId),
      ),
    )
    .orderBy(asc(mediaImageItems.sortOrder));

  const imagesByBookId = new Map<string, MediaImageItem[]>();
  for (const row of rows) {
    const bookId = row.bookId;
    if (!bookId) {
      continue;
    }
    const currentList = imagesByBookId.get(bookId) ?? [];
    currentList.push({
      url: row.path,
      isDefault: row.isDefault,
      spineColor: row.spineColor,
      selected: false,
    });
    imagesByBookId.set(bookId, currentList);
  }
  return imagesByBookId;
}

export async function loadOtherMediaImagesById(db: Db, otherMediaIds: string[]) {
  if (otherMediaIds.length === 0) {
    return new Map<string, MediaImageItem[]>();
  }

  const rows = await db
    .select({
      otherMediaId: mediaImageItems.otherMediaId,
      path: mediaImageItems.path,
      sortOrder: mediaImageItems.sortOrder,
      isDefault: mediaImageItems.isDefault,
      spineColor: mediaImageItems.spineColor,
    })
    .from(mediaImageItems)
    .where(
      and(
        inArray(mediaImageItems.otherMediaId, otherMediaIds),
        isNull(mediaImageItems.bookId),
      ),
    )
    .orderBy(asc(mediaImageItems.sortOrder));

  const imagesByOtherMediaId = new Map<string, MediaImageItem[]>();
  for (const row of rows) {
    const otherMediaId = row.otherMediaId;
    if (!otherMediaId) {
      continue;
    }
    const currentList = imagesByOtherMediaId.get(otherMediaId) ?? [];
    currentList.push({
      url: row.path,
      isDefault: row.isDefault,
      spineColor: row.spineColor,
      selected: false,
    });
    imagesByOtherMediaId.set(otherMediaId, currentList);
  }
  return imagesByOtherMediaId;
}

export async function loadBookImageUrlsById(db: Db, bookIds: string[]) {
  const imagesByBookId = await loadBookImagesById(db, bookIds);
  return new Map(
    [...imagesByBookId.entries()].map(([bookId, images]) => [
      bookId,
      images.map((image) => image.url),
    ]),
  );
}

export async function loadOtherMediaImageUrlsById(db: Db, otherMediaIds: string[]) {
  const imagesByOtherMediaId = await loadOtherMediaImagesById(db, otherMediaIds);
  return new Map(
    [...imagesByOtherMediaId.entries()].map(([otherMediaId, images]) => [
      otherMediaId,
      images.map((image) => image.url),
    ]),
  );
}
