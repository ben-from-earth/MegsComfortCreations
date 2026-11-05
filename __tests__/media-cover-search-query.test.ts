import {
  mediaCoverResultCount,
  mediaCoverSearchQuery,
} from 'lib/trpc/routers/collect/actions/get-media-covers';

describe('mediaCoverSearchQuery', () => {
  test('searches a movie as its full DVD cover', () => {
    expect(
      mediaCoverSearchQuery('The Princess Bride', 'William Goldman', 'movie'),
    ).toBe('The Princess Bride movie full dvd cover');
  });

  test('keeps the author on other media cover searches', () => {
    expect(mediaCoverSearchQuery('Dune', 'Frank Herbert', 'book')).toBe(
      'Dune Frank Herbert book Cover Image',
    );
    expect(mediaCoverSearchQuery('Abbey Road', undefined, 'album')).toBe(
      'Abbey Road album Cover Image',
    );
  });
});

describe('mediaCoverResultCount', () => {
  test('asks for five movie covers and three for other media', () => {
    expect(mediaCoverResultCount('movie')).toBe(5);
    expect(mediaCoverResultCount('book')).toBe(3);
    expect(mediaCoverResultCount('album')).toBe(3);
    expect(mediaCoverResultCount('videoGame')).toBe(3);
  });
});
