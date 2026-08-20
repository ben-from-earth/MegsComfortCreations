import {
  DATABASE_EDIT_FAILED_MESSAGE,
  IMAGE_PERSISTENCE_FAILED_MESSAGE,
  MEDIA_NOT_FOUND_MESSAGE,
  SCHEMA_VIOLATION_MESSAGE,
  toDatabaseEditDisplayError,
} from '@/showdatabase/database-edit-error-display';

describe('toDatabaseEditDisplayError', () => {
  test('maps image persistence to the images field', () => {
    expect(toDatabaseEditDisplayError('Image Persistence Error')).toEqual({
      placement: 'field',
      field: 'images',
      message: IMAGE_PERSISTENCE_FAILED_MESSAGE,
    });
  });

  test('maps missing items, schema fallbacks, and unknown failures to the form banner', () => {
    expect(toDatabaseEditDisplayError('Media Not Found')).toEqual({
      placement: 'form',
      message: MEDIA_NOT_FOUND_MESSAGE,
    });
    expect(toDatabaseEditDisplayError('Schema Violation')).toEqual({
      placement: 'form',
      message: SCHEMA_VIOLATION_MESSAGE,
    });
    expect(toDatabaseEditDisplayError('Unexpected Failure')).toEqual({
      placement: 'form',
      message: DATABASE_EDIT_FAILED_MESSAGE,
    });
  });
});
