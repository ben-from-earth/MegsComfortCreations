import { genresRouter } from '../../lib/trpc/routers/genres/_';
import { loadBookImageUrlsById } from 'lib/media-storage/media-image-records';

jest.mock('lib/media-storage/media-image-records', () => ({
  loadBookImageUrlsById: jest.fn(),
}));

function createAdminCaller(db: unknown) {
  return genresRouter.createCaller({
    db,
    authSession: { user: { id: '1', role: 'admin' } },
    user: { id: '1', role: 'admin' },
  } as never);
}

describe('genres router', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    (loadBookImageUrlsById as jest.Mock).mockResolvedValue(new Map());
  });

  test('getAll returns full genre list', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest
          .fn()
          .mockResolvedValueOnce([
            { genre: 'Fantasy' },
            { genre: 'Science Fiction' },
          ]),
      })),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.getAll();
    expect(response).toEqual({
      message: 'Success',
      genres: ['Fantasy', 'Science Fiction'],
    });
  });

  test('getForBook returns book-linked genres', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          innerJoin: jest.fn(() => ({
            where: jest.fn().mockResolvedValueOnce([{ genre: 'Mystery' }]),
          })),
        })),
      })),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.getForBook({
      bookID: 'f58ca8c2-71b6-4ca8-9335-db8cf797f63d',
    });

    expect(response.genres).toEqual(['Mystery']);
  });

  test('link returns success and not-found responses per genre', async () => {
    const selectWhere = jest
      .fn()
      .mockResolvedValueOnce([{ id: 'genre-1' }])
      .mockResolvedValueOnce([]);
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({ where: selectWhere })),
      })),
      insert: jest.fn(() => ({
        values: jest.fn().mockResolvedValue(undefined),
      })),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.link({
      bookID: 'f58ca8c2-71b6-4ca8-9335-db8cf797f63d',
      genres: ['Fantasy', 'Unknown'],
    });

    expect(response.genreResponses).toEqual([
      {
        message: 'Successful genre link',
        genre: 'Fantasy',
        bookID: 'f58ca8c2-71b6-4ca8-9335-db8cf797f63d',
      },
      {
        message: 'Genre "Unknown" not found in database.',
        genre: 'Unknown',
        bookID: 'f58ca8c2-71b6-4ca8-9335-db8cf797f63d',
      },
    ]);
    expect(mockDb.insert).toHaveBeenCalledTimes(1);
  });

  test('unlink deletes link rows when genres are found', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          where: jest.fn().mockResolvedValueOnce([{ id: 'genre-1' }]),
        })),
      })),
      delete: jest.fn(() => ({
        where: jest.fn().mockResolvedValue(undefined),
      })),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.unlink({
      bookID: 'f58ca8c2-71b6-4ca8-9335-db8cf797f63d',
      genres: ['Fantasy'],
    });

    expect(mockDb.delete).toHaveBeenCalledTimes(1);
    expect(response.genreResponses[0].message).toBe('Successful genre unlink');
  });

  test('paginateByGenre returns paginated books response contract', async () => {
    const mockDb = {
      select: jest
        .fn()
        .mockReturnValueOnce({
          from: jest.fn(() => ({
            innerJoin: jest.fn(() => ({
              innerJoin: jest.fn(() => ({
                where: jest.fn(() => ({
                  orderBy: jest.fn(() => ({
                    limit: jest.fn(() => ({
                      offset: jest
                        .fn()
                        .mockResolvedValueOnce([
                          { book: { id: 'book-1', title: 'Dune' } },
                        ]),
                    })),
                  })),
                })),
              })),
            })),
          })),
        })
        .mockReturnValueOnce({
          from: jest.fn(() => ({
            innerJoin: jest.fn(() => ({
              innerJoin: jest.fn(() => ({
                where: jest.fn().mockResolvedValueOnce([{ value: 1 }]),
              })),
            })),
          })),
        }),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.paginateByGenre({
      genre: 'Science Fiction',
      limit: 10,
      page: 1,
      sort: 'title',
      ascDesc: 'asc',
    });

    expect(response).toEqual({
      message: 'Successful database gather',
      paginatedList: [{ id: 'book-1', title: 'Dune', imageUrls: [] }],
      total: 1,
    });
  });

  test('paginateNoGenres returns books without linked genres', async () => {
    const mockDb = {
      select: jest
        .fn()
        .mockReturnValueOnce({
          from: jest.fn(() => ({
            leftJoin: jest.fn(() => ({
              where: jest.fn(() => ({
                orderBy: jest.fn(() => ({
                  limit: jest.fn(() => ({
                    offset: jest
                      .fn()
                      .mockResolvedValueOnce([
                        { book: { id: 'book-2', title: 'Standalone' } },
                      ]),
                  })),
                })),
              })),
            })),
          })),
        })
        .mockReturnValueOnce({
          from: jest.fn(() => ({
            leftJoin: jest.fn(() => ({
              where: jest.fn().mockResolvedValueOnce([{ value: 1 }]),
            })),
          })),
        }),
    };

    const caller = createAdminCaller(mockDb);
    const response = await caller.paginateNoGenres({
      limit: 10,
      page: 1,
      sort: 'title',
      ascDesc: 'asc',
    });

    expect(response).toEqual({
      message: 'Successful database gather',
      paginatedList: [{ id: 'book-2', title: 'Standalone', imageUrls: [] }],
      total: 1,
    });
  });
});
