import { jest } from '@jest/globals';

describe('online router', () => {
  const OLD_ENV = process.env;

  beforeEach(() => {
    jest.resetModules();
    jest.restoreAllMocks();
    process.env = { ...OLD_ENV };
  });

  afterAll(() => {
    process.env = OLD_ENV;
  });

  test('openLibrary returns friendly error when author is missing', async () => {
    const { onlineRouter } = await import('lib/trpc/routers/online');
    const caller = onlineRouter.createCaller({
      authSession: { user: { id: '1', role: 'admin' } },
      user: { id: '1', role: 'admin' },
      db: {},
    } as never);

    const response = await caller.openLibrary({ title: 'Dune' });
    expect(response).toEqual({
      error: 'Open Library Error',
      message: 'Error gathering Open Library data for Dune, author not provided',
      failedSearchData: { title: 'Dune', author: undefined },
    });
  });

  test('openLibrary returns gathered data with valid response', async () => {
    const axios = (await import('axios')).default;
    jest.spyOn(axios, 'get').mockResolvedValueOnce({
      data: {
        docs: [{ first_publish_year: 1965, number_of_pages_median: 412 }],
      },
    } as never);

    const { onlineRouter } = await import('lib/trpc/routers/online');
    const caller = onlineRouter.createCaller({
      authSession: { user: { id: '1', role: 'admin' } },
      user: { id: '1', role: 'admin' },
      db: {},
    } as never);

    const response = await caller.openLibrary({
      title: 'Dune',
      author: 'Frank Herbert',
    });
    expect(response).toEqual({
      title: 'Dune',
      author: 'Frank Herbert',
      pubYear: 1965,
      pageCount: 412,
    });
  });

  test('mediaCovers returns credential error when env vars are missing', async () => {
    delete process.env.GOOGLE_SEARCH_API_KEY;
    delete process.env.GOOGLE_SEARCH_CX;

    const { onlineRouter } = await import('lib/trpc/routers/online');
    const caller = onlineRouter.createCaller({
      authSession: { user: { id: '1', role: 'admin' } },
      user: { id: '1', role: 'admin' },
      db: {},
    } as never);

    const response = await caller.mediaCovers({ title: 'Dune', type: 'book' });
    expect(response).toEqual({
      error: 'Google Search Credential Error',
      message:
        'Error Connecting to Google Search API because of invalid or empty credentials',
      failedSearchData: [],
    });
  });

  test('mediaCovers returns images when credentials are configured', async () => {
    process.env.GOOGLE_SEARCH_API_KEY = 'test-key';
    process.env.GOOGLE_SEARCH_CX = 'test-cx';
    const axios = (await import('axios')).default;
    jest.spyOn(axios, 'get').mockResolvedValueOnce({
      data: { items: [{ link: 'https://img/1' }, { link: 'https://img/2' }] },
    } as never);

    const { onlineRouter } = await import('lib/trpc/routers/online');
    const caller = onlineRouter.createCaller({
      authSession: { user: { id: '1', role: 'admin' } },
      user: { id: '1', role: 'admin' },
      db: {},
    } as never);

    const response = await caller.mediaCovers({
      title: 'Dune',
      author: 'Frank Herbert',
      type: 'book',
    });
    expect(response).toEqual({ images: ['https://img/1', 'https://img/2'] });
  });
});
