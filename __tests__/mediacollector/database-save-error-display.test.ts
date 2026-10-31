import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  toUserFriendlyDatabaseSaveReason,
} from '@/mediacollector/database-save-error-display';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';
import type { DatabaseSaveFailureResult } from 'lib/interfaces/global-interfaces';

function createBlock(overrides: Partial<MediaItemForm> = {}): MediaItemForm {
  return {
    type: 'book',
    images: [
      {
        url: 'https://img/default.png',
        selected: true,
        isDefault: true,
        spineColor: '#111111',
      },
    ],
    blockInfo: {
      title: 'Dune',
      spineColor: '#111111',
      genres: [],
    },
    blockID: 'BLK-1',
    isDatabase: false,
    ...overrides,
  };
}

function saveFailure(
  overrides: Partial<DatabaseSaveFailureResult> = {},
): DatabaseSaveFailureResult {
  return {
    success: false,
    title: 'Dune',
    error: 'Schema Violation',
    message: 'Schema violation(s) during save request',
    errors: ['Required'],
    blockID: 'BLK-1',
    ...overrides,
  };
}

describe('toUserFriendlyDatabaseSaveReason', () => {
  test.each([
    [
      saveFailure({
        error: 'Image Persistence Error',
        message: 'Image failed to save so media item creation was rolled back.',
        errors: [
          'Image failed to save so media item creation was rolled back.',
        ],
      }),
      'The cover image could not be saved, so this item was not added to the database.',
    ],
    [
      saveFailure({
        error: 'Cover Append Error',
        message: 'The new cover could not be added to this database item.',
        errors: ['Book was not found'],
      }),
      'The new cover could not be added to this database item.',
    ],
    [
      saveFailure(),
      'Some required details for this item were missing or invalid.',
    ],
    [
      saveFailure({
        error: 'Database Insertion Error',
        message: 'An error occurred while trying to save to the database',
        errors: ['Genre "Sci-Fi" does not exist'],
      }),
      'A selected genre is not available in the database.',
    ],
    [
      saveFailure({
        error: 'Database Insertion Error',
        message: 'An error occurred while trying to save to the database',
        errors: ['Book insertion failed'],
      }),
      'This item could not be saved to the database. Try again, or remove the block and re-collect it.',
    ],
    [
      saveFailure({
        error: 'Media Not Found',
        message: 'The database item could not be found.',
        errors: ['Book was not found'],
      }),
      'This database item could not be found, so the edit was not saved.',
    ],
    [
      saveFailure({
        error: 'Database Edit Error',
        message: 'An error occurred while trying to save the edit',
        errors: ['Unknown error'],
      }),
      'The changes to this database item could not be saved.',
    ],
    [
      saveFailure({
        error: 'Genre Update Error',
        message: 'The genres for this item could not be updated',
        errors: ['Unknown error'],
      }),
      'The changes were saved, but the genres could not be updated.',
    ],
  ] as const)('maps $error', (failure, message) => {
    expect(toUserFriendlyDatabaseSaveReason(failure)).toBe(message);
  });
});

describe('buildDatabaseSaveFailureDisplayLines', () => {
  test('includes 1-based block numbers matching collector card order', () => {
    expect(
      buildDatabaseSaveFailureDisplayLines(
        [
          saveFailure({
            blockID: 'BLK-2',
            title: 'Matrix',
            error: 'Image Persistence Error',
            message:
              'Image failed to save so media item creation was rolled back.',
            errors: [
              'Image failed to save so media item creation was rolled back.',
            ],
          }),
        ],
        [
          createBlock(),
          createBlock({
            blockID: 'BLK-2',
            type: 'movie',
            blockInfo: { title: 'Matrix', spineColor: '#222', genres: [] },
          }),
        ],
      ),
    ).toEqual([
      {
        blockID: 'BLK-2',
        title: 'Matrix',
        blockNumber: 2,
        reason:
          'The cover image could not be saved, so this item was not added to the database.',
      },
    ]);
  });
});

describe('markSuccessfulBlocksAsInDatabase', () => {
  test('marks only successful save blockIDs as isDatabase', () => {
    const updated = markSuccessfulBlocksAsInDatabase(
      [
        createBlock({ blockID: 'BLK-1' }),
        createBlock({
          blockID: 'BLK-2',
          type: 'movie',
          blockInfo: { title: 'Matrix', spineColor: '#222', genres: [] },
        }),
        createBlock({
          blockID: 'BLK-3',
          isDatabase: true,
          type: 'album',
          blockInfo: { title: 'Already Saved', spineColor: '#333', genres: [] },
        }),
      ],
      [
        {
          success: true,
          blockID: 'BLK-1',
          title: 'Dune',
          message: 'Dune successfully added to database.',
          type: 'book',
          actionAttemptItem: {
            id: 'book-1',
            title: 'Dune',
            spineColor: '#111111',
            images: [],
          },
        },
        saveFailure({
          blockID: 'BLK-2',
          title: 'Matrix',
          error: 'Image Persistence Error',
        }),
      ],
    );

    expect(updated.map((block) => block.isDatabase)).toEqual([
      true,
      false,
      true,
    ]);
  });
});
