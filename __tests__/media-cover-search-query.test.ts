import {
  mediaCoverResultCount,
  mediaCoverSearchQuery,
} from 'lib/trpc/routers/collect/actions/get-media-covers';

describe('mediaCoverSearchQuery', () => {
  test('searches a movie as its full DVD cover', () => {
    expect(
      mediaCoverSearchQuery('The Princess Bride', 'William Goldman', 'movie'),
    ).toBe('The Princess Bride movie full dvd cover');
    expect(
      mediaCoverSearchQuery(
        'The Princess Bride',
        'William Goldman',
        'movie',
        'dvd',
      ),
    ).toBe('The Princess Bride movie full dvd cover');
  });

  test('searches VHS and Blu-Ray covers by format', () => {
    expect(
      mediaCoverSearchQuery('The Princess Bride', undefined, 'movie', 'vhs'),
    ).toBe('The Princess Bride movie full vhs cover');
    expect(
      mediaCoverSearchQuery(
        'The Princess Bride',
        undefined,
        'movie',
        'bluRay',
      ),
    ).toBe('The Princess Bride movie full blu-ray cover');
  });

  test('keeps the author on book and video game cover searches', () => {
    expect(mediaCoverSearchQuery('Dune', 'Frank Herbert', 'book')).toBe(
      'Dune Frank Herbert book Cover Image',
    );
    expect(
      mediaCoverSearchQuery('The Legend of Zelda', undefined, 'videoGame'),
    ).toBe('The Legend of Zelda videoGame Cover Image');
  });
});

describe('mediaCoverResultCount', () => {
  test('asks for five movie covers and three for books and video games', () => {
    expect(mediaCoverResultCount('movie')).toBe(5);
    expect(mediaCoverResultCount('book')).toBe(3);
    expect(mediaCoverResultCount('videoGame')).toBe(3);
  });
});
