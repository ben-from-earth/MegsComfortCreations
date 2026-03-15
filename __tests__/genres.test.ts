import { genresRouter } from '../lib/trpc/routers/genres/_';

function createAdminCaller(db: unknown) {
  return genresRouter.createCaller({
    db,
    authSession: { user: { id: '1', role: 'admin' } },
    user: { id: '1', role: 'admin' },
  } as never);
}

describe('genres router smoke coverage', () => {
  test('getAll returns genres from current tRPC router', async () => {
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

    expect(response.message).toBe('Success');
    expect(response.genres).toEqual(['Fantasy', 'Science Fiction']);
  });

  test('getForBook returns expected contract for a book id', async () => {
    const mockDb = {
      select: jest.fn(() => ({
        from: jest.fn(() => ({
          innerJoin: jest.fn(() => ({
            where: jest.fn().mockResolvedValueOnce([{ genre: 'Fantasy' }]),
          })),
        })),
      })),
    };

    const caller = createAdminCaller(mockDb);
    const bookID = '7f2be7ec-2d88-4ae0-b4ec-314f7221b7ba';
    const response = await caller.getForBook({ bookID });

    expect(response.message).toBe(`Successfully grabbed genres for bookID ${bookID}`);
    expect(response.genres).toEqual(['Fantasy']);
  });
});
