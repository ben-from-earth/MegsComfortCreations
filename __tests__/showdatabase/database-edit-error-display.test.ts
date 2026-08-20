import {
  DATABASE_EDIT_FAILED_MESSAGE,
  toUserFriendlyDatabaseEditReason,
} from '@/showdatabase/database-edit-error-display';

describe('toUserFriendlyDatabaseEditReason', () => {
  test('maps image persistence errors', () => {
    expect(
      toUserFriendlyDatabaseEditReason({
        error: 'Image Persistence Error',
      }),
    ).toBe('The cover image could not be saved, so the edit was not applied.');
  });

  test('maps schema violations', () => {
    expect(
      toUserFriendlyDatabaseEditReason({
        error: 'Schema Violation',
      }),
    ).toBe('Some required details for this item were missing or invalid.');
  });

  test('maps missing database items', () => {
    expect(
      toUserFriendlyDatabaseEditReason({
        error: 'Media Not Found',
      }),
    ).toBe('This item is no longer in the database.');
  });

  test('maps unknown edit errors', () => {
    expect(
      toUserFriendlyDatabaseEditReason({
        error: 'Unexpected Failure',
      }),
    ).toBe(DATABASE_EDIT_FAILED_MESSAGE);
  });
});
