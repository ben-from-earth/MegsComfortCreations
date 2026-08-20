import type {
  DatabaseSaveFailureResult,
  DatabaseSaveServerResponse,
} from 'lib/interfaces/globalInterfaces';
import type { MediaItemForm } from './collector-form/mediaItemFormSchema';

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

  if (item.error === 'Schema Violation') {
    return 'Some required details for this item were missing or invalid.';
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
