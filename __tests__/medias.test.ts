import { searchByTitle } from 'lib/trpc/routers/database/actions/search-by-title';
import { createAdminContext, createTrpcCaller } from './helpers/trpcTestContext';

jest.mock('lib/trpc/routers/database/actions/search-by-title', () => ({
  searchByTitle: jest.fn(),
}));

describe('database router smoke coverage', () => {
  test('searchByTitle delegates to active tRPC search action', async () => {
    const mockedSearchByTitle = searchByTitle as jest.Mock;
    mockedSearchByTitle.mockResolvedValueOnce({
      message: 'Successfully found 1 book(s) with title Dune',
      foundMediaList: [{ id: 'book-1', title: 'Dune' }],
      total: 1,
    });

    const caller = createTrpcCaller(createAdminContext({}));
    const response = await caller.database.searchByTitle({
      type: 'book',
      title: 'dune',
    });

    expect(response.total).toBe(1);
    expect(response.foundMediaList[0].id).toBe('book-1');
    expect(mockedSearchByTitle).toHaveBeenCalledWith(
      expect.anything(),
      'book',
      'dune',
    );
  });

  test('getPaginated rejects unsupported non-book sort options', async () => {
    const caller = createTrpcCaller(createAdminContext({}));

    await expect(
      caller.database.getPaginated({
        type: 'movie',
        limit: 10,
        page: 1,
        sort: 'author',
        ascDesc: 'asc',
        genre: '',
      }),
    ).rejects.toMatchObject({
      code: 'BAD_REQUEST',
      message: expect.stringContaining('Sort "author" is not supported for movie'),
    });
  });
});
