import {
  DATABASE_EDIT_FAILED_MESSAGE,
  toUserFriendlyDatabaseEditReason,
} from '@/showdatabase/database-edit-error-display';

describe('toUserFriendlyDatabaseEditReason', () => {
  test.each([
    [
      'Image Persistence Error',
      'The cover image could not be saved, so the edit was not applied.',
    ],
    [
      'Schema Violation',
      'Some required details for this item were missing or invalid.',
    ],
    ['Media Not Found', 'This item is no longer in the database.'],
    ['Unexpected Failure', DATABASE_EDIT_FAILED_MESSAGE],
  ] as const)('maps %s', (error, message) => {
    expect(toUserFriendlyDatabaseEditReason({ error })).toBe(message);
  });
});
