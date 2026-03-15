import { collectRouter } from 'lib/trpc/routers/collect/_';
import { getMediaCovers } from 'lib/trpc/routers/collect/actions/get-media-covers';
import { getOpenLibraryData } from 'lib/trpc/routers/collect/actions/get-open-library-data';
import { searchByTitle } from 'lib/trpc/routers/database/actions/search-by-title';

jest.mock('lib/trpc/routers/collect/actions/get-media-covers', () => ({
  getMediaCovers: jest.fn(),
}));

jest.mock('lib/trpc/routers/collect/actions/get-open-library-data', () => ({
  getOpenLibraryData: jest.fn(),
}));

jest.mock('lib/trpc/routers/database/actions/search-by-title', () => ({
  searchByTitle: jest.fn(),
}));

describe('collect router', () => {
  test('collectMedia returns database book block with resolved genres', async () => {
    const mockedSearchByTitle = searchByTitle as jest.Mock;
    mockedSearchByTitle.mockResolvedValueOnce({
      total: 1,
      foundMediaList: [
        {
          id: 'book-id-1',
          title: 'Dune',
          author: 'Frank Herbert',
          pageCount: 412,
          pubYear: 1965,
          spineColor: '#fff',
          imageUrls: ['https://img/book-1.png'],
        },
      ],
    });

    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          innerJoin: jest.fn(() => ({
            where: jest.fn().mockResolvedValue([{ genre: 'Science Fiction' }]),
          })),
        })),
      })),
      insert: jest.fn(() => ({
        values: jest.fn(() => ({
          onConflictDoUpdate: jest.fn().mockResolvedValue(undefined),
        })),
      })),
    };

    const caller = collectRouter.createCaller({
      db: mockDb,
      authSession: { user: { id: '1', role: 'admin' } },
      user: { id: '1', role: 'admin' },
    } as never);

    const response = await caller.collectMedia({
      book: [{ title: 'Dune', author: 'Frank Herbert' }],
      movie: [],
      videoGame: [],
      album: [],
    });

    expect(response).toHaveLength(1);
    expect(response[0]).toMatchObject({
      type: 'book',
      isDatabase: true,
      blockInfo: {
        title: 'Dune',
        author: 'Frank Herbert',
        genres: ['Science Fiction'],
      },
    });
  });

  test('collectMedia collects remote data and increments query usage', async () => {
    const mockedSearchByTitle = searchByTitle as jest.Mock;
    const mockedGetMediaCovers = getMediaCovers as jest.Mock;
    const mockedGetOpenLibraryData = getOpenLibraryData as jest.Mock;

    mockedSearchByTitle.mockResolvedValueOnce({ total: 0, foundMediaList: [] });
    mockedGetMediaCovers.mockResolvedValueOnce([
      'https://img/1.png',
      'https://img/2.png',
      'https://img/3.png',
    ]);
    mockedGetOpenLibraryData.mockResolvedValueOnce({
      title: 'Dune',
      author: 'Frank Herbert',
      pubYear: 1965,
      pageCount: 412,
    });

    const onConflictDoUpdate = jest.fn().mockResolvedValue(undefined);
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          innerJoin: jest.fn(() => ({
            where: jest.fn().mockResolvedValue([]),
          })),
        })),
      })),
      insert: jest.fn(() => ({
        values: jest.fn(() => ({
          onConflictDoUpdate,
        })),
      })),
    };

    const caller = collectRouter.createCaller({
      db: mockDb,
      authSession: { user: { id: '1', role: 'admin' } },
      user: { id: '1', role: 'admin' },
    } as never);

    const response = await caller.collectMedia({
      book: [{ title: 'Dune', author: 'Frank Herbert' }],
      movie: [],
      videoGame: [],
      album: [],
    });

    expect(response[0]).toMatchObject({
      isDatabase: false,
      type: 'book',
      blockInfo: {
        title: 'Dune',
        author: 'Frank Herbert',
        pubYear: 1965,
      },
    });
    expect(onConflictDoUpdate).toHaveBeenCalledTimes(1);
  });
});
