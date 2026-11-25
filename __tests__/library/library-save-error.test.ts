import {
  LIBRARY_SAVE_FAILED_MESSAGE,
  IMAGE_PERSISTENCE_FAILED_MESSAGE,
  MEDIA_NOT_FOUND_MESSAGE,
  SCHEMA_VIOLATION_MESSAGE,
  toLibrarySaveDisplayError,
} from '@/library/library-save-error';

describe('toLibrarySaveDisplayError', () => {
  test('maps image persistence to the images field', () => {
    expect(toLibrarySaveDisplayError('Image Persistence Error')).toEqual({
      placement: 'field',
      field: 'images',
      message: IMAGE_PERSISTENCE_FAILED_MESSAGE,
    });
  });

  test('maps missing items, schema fallbacks, and unknown failures to the form banner', () => {
    expect(toLibrarySaveDisplayError('Media Not Found')).toEqual({
      placement: 'form',
      message: MEDIA_NOT_FOUND_MESSAGE,
    });
    expect(toLibrarySaveDisplayError('Schema Violation')).toEqual({
      placement: 'form',
      message: SCHEMA_VIOLATION_MESSAGE,
    });
    expect(toLibrarySaveDisplayError('Unexpected Failure')).toEqual({
      placement: 'form',
      message: LIBRARY_SAVE_FAILED_MESSAGE,
    });
  });
});
