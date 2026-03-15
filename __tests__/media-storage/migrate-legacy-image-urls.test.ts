import {
  migrateLegacyImageUrlsToLocalFiles,
  runOneTimeLegacyImageMigration,
} from 'lib/media-storage/migrate-legacy-image-urls';
import {
  getImageMigrationStatus,
  replaceBookImageRecords,
  replaceOtherMediaImageRecords,
  resolveAndPersistImageList,
} from 'lib/media-storage/media-image-records';

jest.mock('lib/media-storage/media-image-records', () => ({
  getImageMigrationStatus: jest.fn(),
  replaceBookImageRecords: jest.fn().mockResolvedValue(undefined),
  replaceOtherMediaImageRecords: jest.fn().mockResolvedValue(undefined),
  resolveAndPersistImageList: jest.fn(),
}));

describe('legacy image migration service', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('dry-run reports migration impact without writing', async () => {
    (resolveAndPersistImageList as jest.Mock)
      .mockResolvedValueOnce({
        images: [
          {
            publicPath: '/uploads/covers/2026/03/book.png',
            sourceUrl: 'https://legacy.example.com/book.png',
            mimeType: 'image/png',
            sizeBytes: 1000,
          },
        ],
        failures: [],
      })
      .mockResolvedValueOnce({
        images: [
          {
            publicPath: '/uploads/covers/2026/03/movie.png',
            sourceUrl: 'https://legacy.example.com/movie.png',
            mimeType: 'image/png',
            sizeBytes: 1000,
          },
        ],
        failures: [],
      });

    const db = {
      select: jest
        .fn()
        .mockReturnValueOnce({
          from: jest
            .fn()
            .mockResolvedValue([
              {
                id: 'book-1',
                imageUrls: ['https://legacy.example.com/book.png'],
              },
            ]),
        })
        .mockReturnValueOnce({
          from: jest.fn().mockResolvedValue([
            {
              id: 'movie-1',
              mediaType: 'movie',
              imageUrls: ['https://legacy.example.com/movie.png'],
            },
          ]),
        }),
    };

    const summary = await migrateLegacyImageUrlsToLocalFiles({
      db: db as never,
      dryRun: true,
    });

    expect(summary).toMatchObject({
      dryRun: true,
      processedItems: 2,
      migratedExternalUrls: 2,
      failedDownloads: 0,
      deletedRows: 0,
    });
    expect(replaceBookImageRecords).not.toHaveBeenCalled();
    expect(replaceOtherMediaImageRecords).not.toHaveBeenCalled();
  });

  test('one-time migration short-circuits when already completed', async () => {
    (getImageMigrationStatus as jest.Mock).mockResolvedValue({
      totalItems: 4,
      externalUrlCount: 0,
      missingReferenceCount: 0,
      isCompleted: true,
    });

    const result = await runOneTimeLegacyImageMigration({} as never);

    expect(result).toEqual({
      alreadyCompleted: true,
      statusBefore: {
        totalItems: 4,
        externalUrlCount: 0,
        missingReferenceCount: 0,
        isCompleted: true,
      },
      statusAfter: {
        totalItems: 4,
        externalUrlCount: 0,
        missingReferenceCount: 0,
        isCompleted: true,
      },
      summary: null,
    });
  });

  test('non-dry migration deletes media row when any image fails', async () => {
    (resolveAndPersistImageList as jest.Mock)
      .mockResolvedValueOnce({
        images: [
          {
            publicPath: 'https://legacy.example.com/book-fail.png',
            sourceUrl: 'https://legacy.example.com/book-fail.png',
            mimeType: null,
            sizeBytes: null,
          },
        ],
        failures: [
          {
            sourceUrl: 'https://legacy.example.com/book-fail.png',
            message: 'Download failed',
          },
        ],
      })
      .mockResolvedValueOnce({
        images: [
          {
            publicPath: '/uploads/covers/2026/03/movie.png',
            sourceUrl: 'https://legacy.example.com/movie.png',
            mimeType: 'image/png',
            sizeBytes: 1000,
          },
        ],
        failures: [],
      });

    const db = {
      select: jest
        .fn()
        .mockReturnValueOnce({
          from: jest
            .fn()
            .mockResolvedValue([
              {
                id: 'book-1',
                imageUrls: ['https://legacy.example.com/book-fail.png'],
              },
            ]),
        })
        .mockReturnValueOnce({
          from: jest.fn().mockResolvedValue([
            {
              id: 'movie-1',
              mediaType: 'movie',
              imageUrls: ['https://legacy.example.com/movie.png'],
            },
          ]),
        }),
      delete: jest.fn(() => ({
        where: jest.fn().mockResolvedValue(undefined),
      })),
    };

    const summary = await migrateLegacyImageUrlsToLocalFiles({
      db: db as never,
      dryRun: false,
    });

    expect(summary).toMatchObject({
      failedDownloads: 1,
      deletedRows: 1,
      migratedExternalUrls: 1,
    });
    expect(db.delete).toHaveBeenCalledTimes(1);
    expect(replaceBookImageRecords).not.toHaveBeenCalled();
    expect(replaceOtherMediaImageRecords).toHaveBeenCalledTimes(1);
  });
});
