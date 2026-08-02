import { databaseRouter } from 'lib/trpc/routers/database/_';
import { searchByTitle } from 'lib/trpc/routers/database/actions/search-by-title';
import {
  loadBookImagesById,
  replaceBookImageRecords,
  replaceOtherMediaImageRecords,
  resolveAndPersistImageList,
} from 'lib/media-storage/media-image-records';
import { createAdminContext } from '../helpers/trpcTestContext';
import { createMockDb } from '../helpers/mockDrizzle';

jest.mock('lib/trpc/routers/database/actions/search-by-title', () => ({
  searchByTitle: jest.fn(),
}));
jest.mock('lib/media-storage/media-image-records', () => ({
  loadBookImagesById: jest.fn().mockResolvedValue(new Map()),
  loadOtherMediaImagesById: jest.fn().mockResolvedValue(new Map()),
  loadBookImageUrlsById: jest.fn().mockResolvedValue(new Map()),
  loadOtherMediaImageUrlsById: jest.fn().mockResolvedValue(new Map()),
  replaceBookImageRecords: jest.fn().mockResolvedValue(undefined),
  replaceOtherMediaImageRecords: jest.fn().mockResolvedValue(undefined),
  resolveAndPersistImageList: jest.fn(),
}));

const mockedSearchByTitle = jest.mocked(searchByTitle);
const mockedLoadBookImagesById = jest.mocked(loadBookImagesById);
const mockedResolveAndPersistImageList = jest.mocked(resolveAndPersistImageList);
const mockedReplaceBookImageRecords = jest.mocked(replaceBookImageRecords);
const mockedReplaceOtherMediaImageRecords = jest.mocked(
  replaceOtherMediaImageRecords,
);

const duneBook = {
  id: 'book-1',
  title: 'Dune',
  author: 'Frank Herbert',
  pageCount: 412,
  pubYear: 1965,
  spineColor: '#fff',
} as const;

const matrixMovie = {
  id: 'movie-1',
  mediaType: 'movie' as const,
  title: 'Matrix',
  spineColor: '#ffffff',
};

const duneCoverImage = {
  url: 'https://images.example.com/cover.png',
  isDefault: true,
  spineColor: '#fff',
} as const;

const matrixCoverImage = {
  url: 'https://img/selected.png',
  selected: true,
  isDefault: true,
  spineColor: '#ffffff',
} as const;

function createDatabaseCaller(db: unknown = {}) {
  return databaseRouter.createCaller(createAdminContext(db) as never);
}

function getSourceImageUrl(sourceImage: unknown) {
  if (typeof sourceImage === 'string') {
    return sourceImage;
  }
  if (
    typeof sourceImage !== 'object' ||
    sourceImage === null ||
    !('url' in sourceImage)
  ) {
    return null;
  }
  const url = sourceImage.url;
  return typeof url === 'string' ? url : null;
}

function mockSuccessfulImagePersistence() {
  mockedResolveAndPersistImageList.mockImplementation(
    async (_mediaReference, sourceImages) => ({
      images: sourceImages.flatMap((sourceImage, index) => {
        const sourceUrl = getSourceImageUrl(sourceImage);
        if (!sourceUrl) {
          return [];
        }
        return [
          {
            publicPath: sourceUrl.startsWith('http')
              ? `/uploads/covers/2026/03/${sourceUrl.split('/').pop()}`
              : sourceUrl,
            mimeType: 'image/png',
            sizeBytes: 1024,
            sourceUrl,
            isDefault: index === 0,
            spineColor: '#ffffff',
          },
        ];
      }),
      failures: [],
    }),
  );
}

function mockImagePersistenceFailure(sourceUrl: string) {
  mockedResolveAndPersistImageList.mockResolvedValueOnce({
    images: [],
    failures: [{ sourceUrl, message: 'AccessDenied' }],
  });
}

describe('database router', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    mockedLoadBookImagesById.mockResolvedValue(new Map());
    mockSuccessfulImagePersistence();
  });

  describe('searchByTitle', () => {
    test('delegates to the search action', async () => {
      mockedSearchByTitle.mockResolvedValueOnce({
        message: 'Successfully found 1 book(s) with title Dune',
        foundMediaList: [
          {
            ...duneBook,
            images: [],
          },
        ],
        total: 1,
      });

      const response = await createDatabaseCaller().searchByTitle({
        type: 'book',
        title: duneBook.title,
      });

      expect(response.total).toBe(1);
      expect(mockedSearchByTitle).toHaveBeenCalledWith(
        expect.anything(),
        'book',
        duneBook.title,
      );
    });
  });

  describe('getPaginated', () => {
    test('returns paginated book rows', async () => {
      const rows = [{ ...duneBook, images: [], mediaType: 'book' as const }];
      const mockDb = createMockDb({
        selectResults: [rows, [{ count: 1 }]],
      });

      const response = await createDatabaseCaller(mockDb).getPaginated({
        type: 'book',
        limit: 10,
        page: 1,
        sort: 'title',
        ascDesc: 'asc',
        genre: '',
      });

      expect(response).toEqual({
        message: 'Successful database gather',
        paginatedList: [{ ...rows[0], images: [] }],
        total: 1,
      });
    });

    test('rejects genre filter for non-book media', async () => {
      await expect(
        createDatabaseCaller().getPaginated({
          type: 'movie',
          limit: 10,
          page: 1,
          sort: 'title',
          ascDesc: 'asc',
          genre: 'Fantasy',
        }),
      ).rejects.toMatchObject({
        code: 'BAD_REQUEST',
        message: 'Genre filter is only supported for books',
      });
    });

    test('uses normalized image records when available', async () => {
      const normalizedImages = [
        {
          url: '/uploads/covers/2026/03/book-1.png',
          isDefault: true,
          spineColor: '#fff',
          selected: false,
        },
      ];
      mockedLoadBookImagesById.mockResolvedValueOnce(
        new Map([[duneBook.id, normalizedImages]]),
      );

      const mockDb = createMockDb({
        selectResults: [
          [{ ...duneBook, images: [], mediaType: 'book' as const }],
          [{ count: 1 }],
        ],
      });

      const response = await createDatabaseCaller(mockDb).getPaginated({
        type: 'book',
        limit: 10,
        page: 1,
        sort: 'title',
        ascDesc: 'asc',
        genre: '',
      });

      expect(response.paginatedList[0]?.images).toEqual(normalizedImages);
    });
  });

  describe('delete', () => {
    test('returns not-found message when id does not exist', async () => {
      const mockDb = createMockDb({ deleteResult: [] });

      const response = await createDatabaseCaller(mockDb).delete({
        type: 'book',
        id: 'missing-id',
      });

      expect(response).toEqual({
        message: 'No book item with id: missing-id exists',
      });
    });
  });

  describe('getQueryCount', () => {
    test('returns 0 when no usage row exists for the date', async () => {
      const mockDb = createMockDb({ selectResults: [[]] });

      const response = await createDatabaseCaller(mockDb).getQueryCount({
        date: '2026-03-15',
      });

      expect(response).toEqual({ date: '2026-03-15', queryCount: 0 });
    });
  });

  describe('edit', () => {
    const duneEditItem = {
      ...duneBook,
      images: [duneCoverImage],
    };

    test('returns schema violation payload for invalid book item', async () => {
      const response = await createDatabaseCaller().edit({
        type: 'book',
        item: { title: '', spineColor: '' },
      });

      expect(response).toMatchObject({
        error: 'Schema Violation',
        message: 'Schema violation(s) during edit request',
        type: 'book',
      });
    });

    test('returns media-not-found when the item is missing on select', async () => {
      const mockDb = createMockDb({ selectResults: [[]] });
      const movieEditItem = {
        id: 'missing-id',
        title: 'Nope',
        spineColor: '#000',
        images: [{ url: 'https://img', isDefault: true, spineColor: '#000' }],
      };

      const response = await createDatabaseCaller(mockDb).edit({
        type: 'movie',
        item: movieEditItem,
      });

      expect(response).toEqual({
        error: 'Media Not Found',
        message: 'Edit requested on an item that does not exist in the database',
        actionAttemptItem: movieEditItem,
        type: 'movie',
        errors: ['Nope does not exist in the database.'],
      });
    });

    test('returns media-not-found when book disappears before update', async () => {
      const mockDb = createMockDb({
        selectResults: [[{ ...duneBook, images: [] }]],
        updateResult: [],
      });

      const response = await createDatabaseCaller(mockDb).edit({
        type: 'book',
        item: duneEditItem,
      });

      expect(response).toMatchObject({
        error: 'Media Not Found',
        type: 'book',
        errors: ['Dune does not exist in the database.'],
      });
      expect(mockedReplaceBookImageRecords).not.toHaveBeenCalled();
    });

    test('returns media-not-found when other media disappears before update', async () => {
      const movieEditItem = {
        id: matrixMovie.id,
        title: matrixMovie.title,
        spineColor: '#000',
        images: [{ url: 'https://img', isDefault: true, spineColor: '#000' }],
      };
      const mockDb = createMockDb({
        selectResults: [[{ ...matrixMovie, spineColor: '#000', images: [] }]],
        updateResult: [],
      });

      const response = await createDatabaseCaller(mockDb).edit({
        type: 'movie',
        item: movieEditItem,
      });

      expect(response).toMatchObject({
        error: 'Media Not Found',
        type: 'movie',
        errors: ['Matrix does not exist in the database.'],
      });
      expect(mockedReplaceOtherMediaImageRecords).not.toHaveBeenCalled();
    });

    test('resolves image URLs and rewrites image records for books', async () => {
      const mockDb = createMockDb({
        selectResults: [
          [
            {
              ...duneBook,
              images: [
                {
                  url: '/uploads/covers/2026/03/old-cover.png',
                  isDefault: true,
                  spineColor: '#fff',
                },
              ],
            },
          ],
        ],
        updateResult: [
          {
            ...duneBook,
            images: [
              {
                url: '/uploads/covers/2026/03/cover.png',
                isDefault: true,
                spineColor: '#fff',
              },
            ],
          },
        ],
      });

      const response = await createDatabaseCaller(mockDb).edit({
        type: 'book',
        item: duneEditItem,
      });

      expect(response).toMatchObject({
        type: 'book',
        message: 'Dune successfully edited.',
        actionAttemptItem: {
          id: duneBook.id,
          images: [
            {
              url: '/uploads/covers/2026/03/cover.png',
              isDefault: true,
              spineColor: '#ffffff',
            },
          ],
        },
      });
      expect(mockedResolveAndPersistImageList).toHaveBeenCalled();
      expect(mockedReplaceBookImageRecords).toHaveBeenCalled();
    });

    test('returns friendly image persistence error without updating', async () => {
      mockImagePersistenceFailure(duneCoverImage.url);
      const mockDb = createMockDb({
        selectResults: [
          [
            {
              ...duneBook,
              images: [
                {
                  url: '/uploads/covers/2026/03/old-cover.png',
                  isDefault: true,
                  spineColor: '#fff',
                },
              ],
            },
          ],
        ],
      });

      const response = await createDatabaseCaller(mockDb).edit({
        type: 'book',
        item: duneEditItem,
      });

      expect(response).toEqual({
        title: duneBook.title,
        error: 'Image Persistence Error',
        message: 'Image failed to save so the edit was not applied.',
        errors: ['Image failed to save so the edit was not applied.'],
      });
      expect(mockDb.update).not.toHaveBeenCalled();
      expect(mockedReplaceBookImageRecords).not.toHaveBeenCalled();
    });
  });

  describe('save', () => {
    const matrixSaveInput = {
      type: 'movie' as const,
      blockID: 'BLK-1',
      isDatabase: false,
      images: [
        matrixCoverImage,
        {
          url: 'https://img/unselected.png',
          selected: false,
          isDefault: false,
          spineColor: '#ffffff',
        },
      ],
      blockInfo: {
        title: matrixMovie.title,
        spineColor: matrixMovie.spineColor,
        genres: [] as string[],
      },
    };

    test('persists selected images for new other-media items', async () => {
      const mockDb = createMockDb({
        insertResult: [{ ...matrixMovie, images: [] }],
      });

      const response = await createDatabaseCaller(mockDb).save([matrixSaveInput]);

      expect(response).toEqual([
        {
          message: 'Matrix successfully added to database.',
          actionAttemptItem: {
            ...matrixMovie,
            images: [
              {
                url: '/uploads/covers/2026/03/selected.png',
                isDefault: true,
                spineColor: '#ffffff',
                selected: false,
              },
            ],
            blockID: 'BLK-1',
          },
          type: 'movie',
        },
      ]);
      expect(mockedResolveAndPersistImageList).toHaveBeenCalled();
      expect(mockedReplaceOtherMediaImageRecords).toHaveBeenCalled();
      expect(mockDb.transaction).toHaveBeenCalled();
    });

    test('rolls back the inserted row and returns a friendly S3 error', async () => {
      mockImagePersistenceFailure(matrixCoverImage.url);
      const mockDb = createMockDb({
        insertResult: [{ ...matrixMovie, images: [] }],
      });

      const response = await createDatabaseCaller(mockDb).save([
        {
          ...matrixSaveInput,
          images: [matrixCoverImage],
        },
      ]);

      expect(response).toEqual([
        {
          title: matrixMovie.title,
          error: 'Image Persistence Error',
          message:
            'Image failed to save so media item creation was rolled back.',
          errors: [
            'Image failed to save so media item creation was rolled back.',
          ],
        },
      ]);
      expect(mockDb.delete).toHaveBeenCalledTimes(1);
      expect(mockedReplaceOtherMediaImageRecords).not.toHaveBeenCalled();
    });
  });
});
