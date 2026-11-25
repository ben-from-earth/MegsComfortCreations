import type {
  DatabaseSaveFailureResult,
  DatabaseSaveServerResponse,
  MediaImageItem,
} from 'lib/interfaces/global-interfaces';
import { ORDER_COVER_IN_USE_ERROR } from 'lib/orders/order-cover-in-use';
import type { MediaItemForm } from './collector-form/media-item-form-schema';

export type DatabaseSaveFailureDisplayLine = {
  blockID: string;
  title: string;
  blockNumber: number | null;
  reason: string;
};

export function toUserFriendlyDatabaseSaveReason(
  item: DatabaseSaveFailureResult,
): string {
  if (item.error === 'Image Persistence Error') {
    return 'The cover image could not be saved, so this item was not added to the database.';
  }

  if (item.error === 'Cover Append Error') {
    return 'The new cover could not be added to this database item.';
  }

  if (item.error === 'Media Not Found') {
    return 'This database item could not be found, so the edit was not saved.';
  }

  if (item.error === ORDER_COVER_IN_USE_ERROR) {
    return item.message;
  }

  if (item.error === 'Database Edit Error') {
    return 'The changes to this database item could not be saved.';
  }

  if (item.error === 'Genre Update Error') {
    return 'The changes were saved, but the genres could not be updated.';
  }

  if (item.error === 'Schema Violation') {
    return 'Some required details for this item were missing or invalid.';
  }

  if (item.error.startsWith('Duplicate ') && item.error !== 'Duplicate Book') {
    return item.message;
  }

  if (item.error === 'Database Insertion Error') {
    const hasMissingGenre = item.errors.some(
      (error) => error.includes('Genre "') && error.includes('does not exist'),
    );
    if (hasMissingGenre) {
      return 'A selected genre is not available in the database.';
    }
    return 'This item could not be saved to the database. Try again, or remove the block and re-collect it.';
  }

  return 'This item could not be saved to the database.';
}

export function buildDatabaseSaveFailureDisplayLines(
  saveResults: DatabaseSaveServerResponse,
  collectedData: MediaItemForm[],
): DatabaseSaveFailureDisplayLine[] {
  return saveResults
    .filter((item): item is DatabaseSaveFailureResult => !item.success)
    .map((item) => {
      const blockIndex = collectedData.findIndex(
        (block) => block.blockID === item.blockID,
      );

      return {
        blockID: item.blockID,
        title: item.title,
        blockNumber: blockIndex >= 0 ? blockIndex + 1 : null,
        reason: toUserFriendlyDatabaseSaveReason(item),
      };
    });
}

function copyIds(
  images: MediaItemForm['images'],
  savedImages: readonly MediaImageItem[],
  isDatabase: boolean,
): MediaItemForm['images'] {
  if (!isDatabase) {
    // A new item stores only its selected covers, or the first cover when none
    // is selected, in that order.
    const selectedIndexes = images.flatMap((image, index) =>
      image.selected ? [index] : [],
    );
    const storedIndexes =
      selectedIndexes.length > 0 ? selectedIndexes : images.length > 0 ? [0] : [];
    const savedByFormIndex = new Map(
      storedIndexes.map((formIndex, savedIndex) => [
        formIndex,
        savedImages[savedIndex],
      ]),
    );
    return images.map((image, index) => {
      const saved = savedByFormIndex.get(index);
      return saved?.id ? { ...image, id: saved.id, url: saved.url } : image;
    });
  }

  const savedByUrl = new Map(savedImages.map((image) => [image.url, image]));
  const sameLength = savedImages.length === images.length;
  return images.map((image, index) => {
    const byUrl = savedByUrl.get(image.url);
    if (byUrl?.id) {
      return { ...image, id: byUrl.id };
    }
    const byPosition = sameLength ? savedImages[index] : undefined;
    return byPosition?.id
      ? { ...image, id: byPosition.id, url: byPosition.url }
      : image;
  });
}

export function applySavedCoverIds(
  collectedData: MediaItemForm[],
  saveResults: DatabaseSaveServerResponse,
): MediaItemForm[] {
  const savedImagesByBlockId = new Map(
    saveResults.flatMap((item) =>
      item.success ? [[item.blockID, item.actionAttemptItem.images] as const] : [],
    ),
  );
  if (savedImagesByBlockId.size === 0) {
    return collectedData;
  }

  return collectedData.map((block) => {
    const savedImages = savedImagesByBlockId.get(block.blockID);
    return savedImages
      ? { ...block, images: copyIds(block.images, savedImages, block.isDatabase) }
      : block;
  });
}

export function markSuccessfulBlocksAsInDatabase(
  collectedData: MediaItemForm[],
  saveResults: DatabaseSaveServerResponse,
): MediaItemForm[] {
  const successfulBlockIds = new Set(
    saveResults.filter((item) => item.success).map((item) => item.blockID),
  );

  if (successfulBlockIds.size === 0) {
    return collectedData;
  }

  return collectedData.map((block) =>
    successfulBlockIds.has(block.blockID)
      ? { ...block, isDatabase: true }
      : block,
  );
}
