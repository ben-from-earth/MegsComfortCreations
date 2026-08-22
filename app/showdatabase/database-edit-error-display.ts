export const GENRE_UPDATE_FAILED_MESSAGE =
  'The item was saved, but updating genres failed. Close and try again.';

export const DATABASE_EDIT_FAILED_MESSAGE =
  'This item could not be saved. Please try again.';

export const IMAGE_PERSISTENCE_FAILED_MESSAGE =
  'The cover image could not be saved, so the edit was not applied.';

export const SCHEMA_VIOLATION_MESSAGE =
  'Some required details for this item were missing or invalid.';

export const MEDIA_NOT_FOUND_MESSAGE =
  'This item is no longer in the database.';

export type DatabaseEditFieldName = 'images';

export type DatabaseEditDisplayError =
  | { placement: 'field'; field: DatabaseEditFieldName; message: string }
  | { placement: 'form'; message: string };

export function toDatabaseEditDisplayError(
  error: string,
): DatabaseEditDisplayError {
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

  if (error === 'Schema Violation') {
    return { placement: 'form', message: SCHEMA_VIOLATION_MESSAGE };
  }

  return { placement: 'form', message: DATABASE_EDIT_FAILED_MESSAGE };
}
