import { databaseRouter } from 'lib/trpc/routers/database/_';
import { searchByTitle } from 'lib/trpc/routers/database/actions/search-by-title';

jest.mock('lib/trpc/routers/database/actions/search-by-title', () => ({
  searchByTitle: jest.fn(),
}));

function createAdminCaller(db: unknown) {
  return databaseRouter.createCaller({
    db,
    authSession: { user: { id: '1', role: 'admin' } },
    user: { id: '1', role: 'admin' },
  } as never);
}

describe('database router', () => {
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
      paginatedList: rows,
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
      update: jest.fn(() => ({
        set: jest.fn(() => ({
          where: jest.fn(() => ({
            returning: jest.fn().mockResolvedValue([]),
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
              imageUrls: ['https://img/selected.png'],
            },
          ]),
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
          imageUrls: ['https://img/selected.png'],
          blockID: 'BLK-1',
        },
        type: 'movie',
      },
    ]);
  });
});
