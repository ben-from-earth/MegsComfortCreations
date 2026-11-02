import {
  filterGenresBySearch,
  splitVisibleGenres,
} from '@/components/shared/genre-options';

describe('splitVisibleGenres', () => {
  test('shows the first two and counts the rest', () => {
    expect(splitVisibleGenres(['Romance', 'Fantasy', 'Mystery'])).toEqual({
      visibleGenres: ['Romance', 'Fantasy'],
      moreCount: 1,
    });
    expect(
      splitVisibleGenres(['Romance', 'Fantasy', 'Mystery', 'Horror', 'Memoir']),
    ).toEqual({
      visibleGenres: ['Romance', 'Fantasy'],
      moreCount: 3,
    });
  });

  test('has no overflow at two or fewer', () => {
    expect(splitVisibleGenres(['Romance', 'Fantasy'])).toEqual({
      visibleGenres: ['Romance', 'Fantasy'],
      moreCount: 0,
    });
    expect(splitVisibleGenres([])).toEqual({ visibleGenres: [], moreCount: 0 });
  });
});

describe('filterGenresBySearch', () => {
  const allGenres = ['Romance', 'Romantasy', 'Fantasy', 'Mystery'];

  test('lists every genre when the search is empty', () => {
    expect(filterGenresBySearch(allGenres, '  ')).toEqual(allGenres);
  });

  test('matches the search text as a case-insensitive substring', () => {
    expect(filterGenresBySearch(allGenres, 'TAS')).toEqual([
      'Romantasy',
      'Fantasy',
    ]);
    expect(filterGenresBySearch(allGenres, 'zzz')).toEqual([]);
  });
});
