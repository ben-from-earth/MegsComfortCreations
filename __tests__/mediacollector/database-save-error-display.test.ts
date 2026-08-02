import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  toUserFriendlyDatabaseSaveReason,
} from '@/mediacollector/database-save-error-display';
import type { CollectorFormData } from '@/mediacollector/collector-form/collectorFormSchema';
import type { DatabaseSaveServerResponse } from 'lib/interfaces/globalInterfaces';

type CollectedBlock = CollectorFormData['collectedData'][number];

function createBlock(overrides: Partial<CollectedBlock> = {}): CollectedBlock {
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
      author: 'Frank Herbert',
      pubYear: 1965,
      pageCount: 412,
      spineColor: '#111111',
      genres: [],
    },
    blockID: 'BLK-1',
    isDatabase: false,
    ...overrides,
  };
}

describe('toUserFriendlyDatabaseSaveReason', () => {
  test('maps image persistence errors', () => {
    expect(
      toUserFriendlyDatabaseSaveReason({
        success: false,
        title: 'Dune',
        error: 'Image Persistence Error',
        message: 'Image failed to save so media item creation was rolled back.',
        errors: ['Image failed to save so media item creation was rolled back.'],
        blockID: 'BLK-1',
      }),
    ).toBe(
      'The cover image could not be saved, so this item was not added to the database.',
    );
  });

  test('maps schema violations', () => {
    expect(
      toUserFriendlyDatabaseSaveReason({
        success: false,
        title: 'Dune',
        error: 'Schema Violation',
        message: 'Schema violation(s) during save request',
        errors: ['Required'],
        blockID: 'BLK-1',
      }),
    ).toBe('Some required details for this item were missing or invalid.');
  });

  test('maps missing genre insertion errors', () => {
    expect(
      toUserFriendlyDatabaseSaveReason({
        success: false,
        title: 'Dune',
        error: 'Database Insertion Error',
        message: 'An error occurred while trying to save to the database',
        errors: ['Genre "Sci-Fi" does not exist'],
        blockID: 'BLK-1',
      }),
    ).toBe('A selected genre is not available in the database.');
  });

  test('maps other insertion errors', () => {
    expect(
      toUserFriendlyDatabaseSaveReason({
        success: false,
        title: 'Dune',
        error: 'Database Insertion Error',
        message: 'An error occurred while trying to save to the database',
        errors: ['Book insertion failed'],
        blockID: 'BLK-1',
      }),
    ).toBe(
      'This item could not be saved to the database. Try again, or remove the block and re-collect it.',
    );
  });
});

describe('buildDatabaseSaveFailureDisplayLines', () => {
  test('includes 1-based block numbers matching collector card order', () => {
    const collectedData = [
      createBlock({
        blockID: 'BLK-1',
        blockInfo: { title: 'Dune', spineColor: '#111', genres: [] },
      }),
      createBlock({
        blockID: 'BLK-2',
        type: 'movie',
        blockInfo: { title: 'Matrix', spineColor: '#222', genres: [] },
      }),
    ];
    const saveResults: DatabaseSaveServerResponse = [
      {
        success: false,
        title: 'Matrix',
        error: 'Image Persistence Error',
        message: 'Image failed to save so media item creation was rolled back.',
        errors: ['Image failed to save so media item creation was rolled back.'],
        blockID: 'BLK-2',
      },
    ];

    expect(
      buildDatabaseSaveFailureDisplayLines(saveResults, collectedData),
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
    const collectedData = [
      createBlock({ blockID: 'BLK-1', isDatabase: false }),
      createBlock({
        blockID: 'BLK-2',
        isDatabase: false,
        type: 'movie',
        blockInfo: { title: 'Matrix', spineColor: '#222', genres: [] },
      }),
      createBlock({
        blockID: 'BLK-3',
        isDatabase: true,
        type: 'album',
        blockInfo: { title: 'Already Saved', spineColor: '#333', genres: [] },
      }),
    ];
    const saveResults: DatabaseSaveServerResponse = [
      {
        success: true,
        blockID: 'BLK-1',
        title: 'Dune',
        message: 'Dune successfully added to database.',
        actionAttemptItem: {
          id: 'book-1',
          title: 'Dune',
          spineColor: '#111111',
          images: [],
          blockID: 'BLK-1',
        },
        type: 'book',
      },
      {
        success: false,
        title: 'Matrix',
        error: 'Image Persistence Error',
        message: 'Image failed to save so media item creation was rolled back.',
        errors: ['Image failed to save so media item creation was rolled back.'],
        blockID: 'BLK-2',
      },
    ];

    const updated = markSuccessfulBlocksAsInDatabase(collectedData, saveResults);

    expect(updated[0]?.isDatabase).toBe(true);
    expect(updated[1]?.isDatabase).toBe(false);
    expect(updated[2]?.isDatabase).toBe(true);
  });
});
