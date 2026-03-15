import { databaseRouter } from 'lib/trpc/routers/database/_';
import { searchByTitle } from 'lib/trpc/routers/database/actions/search-by-title';
import {
  loadBookImageUrlsById,
  replaceBookImageRecords,
  replaceOtherMediaImageRecords,
  resolveAndPersistImageList,
} from 'lib/media-storage/media-image-records';

jest.mock('lib/trpc/routers/database/actions/search-by-title', () => ({
  searchByTitle: jest.fn(),
}));
jest.mock('lib/media-storage/media-image-records', () => ({
  loadBookImageUrlsById: jest.fn().mockResolvedValue(new Map()),
  loadOtherMediaImageUrlsById: jest.fn().mockResolvedValue(new Map()),
  replaceBookImageRecords: jest.fn().mockResolvedValue(undefined),
  replaceOtherMediaImageRecords: jest.fn().mockResolvedValue(undefined),
  resolveAndPersistImageList: jest.fn(),
}));

function createAdminCaller(db: unknown) {
  return databaseRouter.createCaller({
    db,
    authSession: { user: { id: '1', role: 'admin' } },
    user: { id: '1', role: 'admin' },
  } as never);
}

describe('database router', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    (resolveAndPersistImageList as jest.Mock).mockImplementation(
      async (_mediaReference: unknown, sourceImageUrls: string[]) => ({
        images: sourceImageUrls.map((sourceUrl) => ({
          publicPath: sourceUrl.startsWith('http')
            ? `/uploads/covers/2026/03/${sourceUrl.split('/').pop()}`
            : sourceUrl,
          mimeType: 'image/png',
          sizeBytes: 1024,
          sourceUrl,
        })),
        failures: [],
      }),
    );
  });

  test('searchByTitle delegates to search action', async () => {
    const mockedSearchByTitle = searchByTitle as jest.Mock;
    mockedSearchByTitle.mockResolvedValueOnce({
      message: 'Successfully found 1 book(s) with title Dune',
      foundMediaList: [{ id: 'book-1', title: 'Dune' }],
      total: 1,
    });

    const caller = createAdminCaller({});
    const response = await caller.searchByTitle({ type: 'book', title: 'Dune' });

    expect(response.total).toBe(1);
    expect(mockedSearchByTitle).toHaveBeenCalledWith(
      expect.anything(),
      'book',
      'Dune',
    );
  });

  test('getPaginated returns paginated book rows', async () => {
    const rows = [
      {
        id: 'book-1',
        title: 'Dune',
        author: 'Frank Herbert',
        pageCount: 412,
        pubYear: 1965,
        spineColor: '#fff',
        imageUrls: ['https://img'],
        mediaType: 'book',
      },
    ];

    const rowsChain = {
      from: jest.fn(() => ({
        leftJoin: jest.fn(() => ({
          leftJoin: jest.fn(() => ({
            where: jest.fn(() => ({
              orderBy: jest.fn(() => ({
                limit: jest.fn(() => ({
                  offset: jest.fn().mockResolvedValue(rows),
                })),
              })),
            })),
          })),
        })),
      })),
    };

    const countChain = {
      from: jest.fn(() => ({
        leftJoin: jest.fn(() => ({
          leftJoin: jest.fn(() => ({
            where: jest.fn().mockResolvedValue([{ count: 1 }]),
          })),
        })),
      })),
    };

    const mockDb = {
      select: jest.fn().mockReturnValueOnce(rowsChain).mockReturnValueOnce(countChain),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.getPaginated({
      type: 'book',
      limit: 10,
      page: 1,
      sort: 'title',
      ascDesc: 'asc',
      genre: '',
    });

    expect(response).toEqual({
      message: 'Successful database gather',
      paginatedList: [
        {
          ...rows[0],
          imageUrls: [],
        },
      ],
      total: 1,
    });
  });

  test('getPaginated rejects genre filter for non-book media', async () => {
    const caller = createAdminCaller({});

    await expect(
      caller.getPaginated({
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

  test('delete returns not-found message when id does not exist', async () => {
    const mockDb = {
      delete: jest.fn(() => ({
        where: jest.fn(() => ({
          returning: jest.fn().mockResolvedValue([]),
        })),
      })),
    };
    const caller = createAdminCaller(mockDb);

    const response = await caller.delete({ type: 'book', id: 'missing-id' });
    expect(response).toEqual({ message: 'No book item with id: missing-id exists' });
  });

  test('getQueryCount returns 0 for missing date row', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          where: jest.fn(() => ({
            limit: jest.fn().mockResolvedValue([]),
          })),
        })),
      })),
    };
    const caller = createAdminCaller(mockDb);

    const response = await caller.getQueryCount({ date: '2026-03-15' });
    expect(response).toEqual({ date: '2026-03-15', queryCount: 0 });
  });

  test('edit returns schema violation payload for invalid book item', async () => {
    const caller = createAdminCaller({});

    const response = await caller.edit({
      type: 'book',
      item: { title: '', spineColor: '' },
    });

    expect(response).toMatchObject({
      error: 'Schema Violation',
      message: 'Schema violation(s) during edit request',
      type: 'book',
    });
  });

  test('edit returns media-not-found for missing non-book item', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          where: jest.fn(() => ({
            limit: jest.fn().mockResolvedValue([]),
          })),
        })),
      })),
    };
    const caller = createAdminCaller(mockDb);

    const response = await caller.edit({
      type: 'movie',
      item: {
        id: 'missing-id',
        title: 'Nope',
        spineColor: '#000',
        imageUrls: ['https://img'],
      },
    });

    expect(response).toEqual({
      error: 'Media Not Found',
      message: 'Edit requested on an item that does not exist in the database',
      actionAttemptItem: {
        id: 'missing-id',
        title: 'Nope',
        spineColor: '#000',
        imageUrls: ['https://img'],
      },
      type: 'movie',
      errors: ['Nope does not exist in the database.'],
    });
  });

  test('save writes non-database, selected-image other media records', async () => {
    const mockDb = {
      insert: jest.fn(() => ({
        values: jest.fn(() => ({
          returning: jest.fn().mockResolvedValue([
            {
              id: 'movie-1',
              mediaType: 'movie',
              title: 'Matrix',
              spineColor: '#ffffff',
              imageUrls: [],
            },
          ]),
        })),
      })),
      update: jest.fn(() => ({
        set: jest.fn(() => ({
          where: jest.fn(() => ({
            returning: jest.fn().mockResolvedValue([
              {
                id: 'movie-1',
                mediaType: 'movie',
                title: 'Matrix',
                spineColor: '#ffffff',
                imageUrls: ['/uploads/covers/2026/03/selected.png'],
              },
            ]),
          })),
        })),
      })),
      transaction: jest.fn(),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.save([
      {
        type: 'movie',
        blockID: 'BLK-1',
        isDatabase: false,
        images: [
          { url: 'https://img/selected.png', selected: true },
          { url: 'https://img/unselected.png', selected: false },
        ],
        blockInfo: {
          title: 'Matrix',
          spineColor: '#ffffff',
          genres: [],
        },
      },
    ]);

    expect(response).toEqual([
      {
        message: 'Matrix successfully added to database.',
        actionAttemptItem: {
          id: 'movie-1',
          mediaType: 'movie',
          title: 'Matrix',
          spineColor: '#ffffff',
          imageUrls: ['/uploads/covers/2026/03/selected.png'],
          blockID: 'BLK-1',
        },
        type: 'movie',
      },
    ]);
    expect(resolveAndPersistImageList).toHaveBeenCalled();
    expect(replaceOtherMediaImageRecords).toHaveBeenCalled();
  });

  test('edit book resolves image URLs and rewrites image records', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          where: jest.fn(() => ({
            limit: jest.fn().mockResolvedValue([
              {
                id: 'book-1',
                title: 'Dune',
                author: 'Frank Herbert',
                pageCount: 412,
                pubYear: 1965,
                spineColor: '#fff',
                imageUrls: ['/uploads/covers/2026/03/old-cover.png'],
              },
            ]),
          })),
        })),
      })),
      update: jest
        .fn(() => ({
          set: jest.fn(() => ({
            where: jest.fn(() => ({
              returning: jest.fn().mockResolvedValue([
                {
                  id: 'book-1',
                  title: 'Dune',
                  author: 'Frank Herbert',
                  pageCount: 412,
                  pubYear: 1965,
                  spineColor: '#fff',
                  imageUrls: ['/uploads/covers/2026/03/cover.png'],
                },
              ]),
            })),
          })),
        })),
    };
    const caller = createAdminCaller(mockDb);

    const response = await caller.edit({
      type: 'book',
      item: {
        id: 'book-1',
        title: 'Dune',
        author: 'Frank Herbert',
        pageCount: 412,
        pubYear: 1965,
        spineColor: '#fff',
        imageUrls: ['https://images.example.com/cover.png'],
      },
    });

    expect(response).toMatchObject({
      type: 'book',
      message: 'Dune successfully edited.',
      actionAttemptItem: {
        id: 'book-1',
        imageUrls: ['/uploads/covers/2026/03/cover.png'],
      },
    });
    expect(resolveAndPersistImageList).toHaveBeenCalled();
    expect(replaceBookImageRecords).toHaveBeenCalled();
  });

  test('save returns image persistence error when S3 upload fails', async () => {
    (resolveAndPersistImageList as jest.Mock).mockResolvedValueOnce({
      images: [],
      failures: [
        {
          sourceUrl: 'https://img/selected.png',
          message: 'AccessDenied',
        },
      ],
    });

    const mockDb = {
      insert: jest.fn(() => ({
        values: jest.fn(() => ({
          returning: jest.fn().mockResolvedValue([
            {
              id: 'movie-1',
              mediaType: 'movie',
              title: 'Matrix',
              spineColor: '#ffffff',
              imageUrls: [],
            },
          ]),
        })),
      })),
      delete: jest.fn(() => ({
        where: jest.fn().mockResolvedValue(undefined),
      })),
      update: jest.fn(),
      transaction: jest.fn(),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.save([
      {
        type: 'movie',
        blockID: 'BLK-1',
        isDatabase: false,
        images: [{ url: 'https://img/selected.png', selected: true }],
        blockInfo: {
          title: 'Matrix',
          spineColor: '#ffffff',
          genres: [],
        },
      },
    ]);

    expect(response).toEqual([
      {
        title: 'Matrix',
        error: 'Image Persistence Error',
        message: 'Failed to persist one or more selected images to S3.',
        errors: [
          'Failed to persist "https://img/selected.png" to S3: AccessDenied',
        ],
      },
    ]);
    expect(mockDb.delete).toHaveBeenCalledTimes(1);
    expect(replaceOtherMediaImageRecords).not.toHaveBeenCalled();
  });

  test('edit returns image persistence error when S3 upload fails', async () => {
    (resolveAndPersistImageList as jest.Mock).mockResolvedValueOnce({
      images: [],
      failures: [
        {
          sourceUrl: 'https://images.example.com/cover.png',
          message: 'AccessDenied',
        },
      ],
    });

    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          where: jest.fn(() => ({
            limit: jest.fn().mockResolvedValue([
              {
                id: 'book-1',
                title: 'Dune',
                author: 'Frank Herbert',
                pageCount: 412,
                pubYear: 1965,
                spineColor: '#fff',
                imageUrls: ['/uploads/covers/2026/03/old-cover.png'],
              },
            ]),
          })),
        })),
      })),
      update: jest.fn(),
    };
    const caller = createAdminCaller(mockDb);

    const response = await caller.edit({
      type: 'book',
      item: {
        id: 'book-1',
        title: 'Dune',
        author: 'Frank Herbert',
        pageCount: 412,
        pubYear: 1965,
        spineColor: '#fff',
        imageUrls: ['https://images.example.com/cover.png'],
      },
    });

    expect(response).toEqual({
      title: 'Dune',
      error: 'Image Persistence Error',
      message: 'Failed to persist one or more selected images to S3.',
      errors: [
        'Failed to persist "https://images.example.com/cover.png" to S3: AccessDenied',
      ],
    });
    expect(mockDb.update).not.toHaveBeenCalled();
    expect(replaceBookImageRecords).not.toHaveBeenCalled();
  });

  test('getPaginated uses normalized image records when available', async () => {
    (loadBookImageUrlsById as jest.Mock).mockResolvedValueOnce(
      new Map([['book-1', ['/uploads/covers/2026/03/book-1.png']]]),
    );

    const rows = [
      {
        id: 'book-1',
        title: 'Dune',
        author: 'Frank Herbert',
        pageCount: 412,
        pubYear: 1965,
        spineColor: '#fff',
        imageUrls: ['https://legacy-url.example.com/image.png'],
        mediaType: 'book',
      },
    ];

    const rowsChain = {
      from: jest.fn(() => ({
        leftJoin: jest.fn(() => ({
          leftJoin: jest.fn(() => ({
            where: jest.fn(() => ({
              orderBy: jest.fn(() => ({
                limit: jest.fn(() => ({
                  offset: jest.fn().mockResolvedValue(rows),
                })),
              })),
            })),
          })),
        })),
      })),
    };

    const countChain = {
      from: jest.fn(() => ({
        leftJoin: jest.fn(() => ({
          leftJoin: jest.fn(() => ({
            where: jest.fn().mockResolvedValue([{ count: 1 }]),
          })),
        })),
      })),
    };

    const mockDb = {
      select: jest
        .fn()
        .mockReturnValueOnce(rowsChain)
        .mockReturnValueOnce(countChain),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.getPaginated({
      type: 'book',
      limit: 10,
      page: 1,
      sort: 'title',
      ascDesc: 'asc',
      genre: '',
    });

    expect(response.paginatedList[0]?.imageUrls).toEqual([
      '/uploads/covers/2026/03/book-1.png',
    ]);
  });
});
