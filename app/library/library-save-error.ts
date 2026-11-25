import {
  ORDER_COVER_IN_USE_ERROR,
  ORDER_COVER_IN_USE_MESSAGE,
} from 'lib/orders/order-cover-in-use';

export const GENRE_UPDATE_FAILED_MESSAGE =
  'The item was saved, but updating genres failed. Save again to retry.';

export const LIBRARY_SAVE_FAILED_MESSAGE =
  'This item could not be saved. Please try again.';

export const IMAGE_PERSISTENCE_FAILED_MESSAGE =
  'The cover image could not be saved, so the edit was not applied.';

export const SCHEMA_VIOLATION_MESSAGE =
  'Some required details for this item were missing or invalid.';

export const MEDIA_NOT_FOUND_MESSAGE =
  'This item is no longer in the database.';

export type LibrarySaveFieldName = 'images';

export type LibrarySaveDisplayError =
  | { placement: 'field'; field: LibrarySaveFieldName; message: string }
  | { placement: 'form'; message: string };

export function toLibrarySaveDisplayError(
  error: string,
): LibrarySaveDisplayError {
  if (error === 'Image Persistence Error') {
    return {
      placement: 'field',
      field: 'images',
      message: IMAGE_PERSISTENCE_FAILED_MESSAGE,
    };
  }

  if (error === 'Media Not Found') {
    return { placement: 'form', message: MEDIA_NOT_FOUND_MESSAGE };
  }

  if (error === ORDER_COVER_IN_USE_ERROR) {
    return { placement: 'form', message: ORDER_COVER_IN_USE_MESSAGE };
  }

  if (error === 'Schema Violation') {
    return { placement: 'form', message: SCHEMA_VIOLATION_MESSAGE };
  }

  return { placement: 'form', message: LIBRARY_SAVE_FAILED_MESSAGE };
}
