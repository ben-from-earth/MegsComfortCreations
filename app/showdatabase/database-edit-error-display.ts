export const GENRE_UPDATE_FAILED_MESSAGE =
  'The item was saved, but updating genres failed. Close and try again.';

export const DATABASE_EDIT_FAILED_MESSAGE =
  'This item could not be saved. Please try again.';

export function toUserFriendlyDatabaseEditReason(item: {
  error: string;
}): string {
  if (item.error === 'Image Persistence Error') {
    return 'The cover image could not be saved, so the edit was not applied.';
  }

  if (item.error === 'Schema Violation') {
    return 'Some required details for this item were missing or invalid.';
  }

  if (item.error === 'Media Not Found') {
    return 'This item is no longer in the database.';
  }

  return DATABASE_EDIT_FAILED_MESSAGE;
}
