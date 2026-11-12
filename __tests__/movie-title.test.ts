import {
  splitMovieTitle,
  toDisplayMovieTitle,
  toStoredMovieTitle,
} from 'lib/helpers/movie-title';

describe('splitMovieTitle', () => {
  test.each([
    ['Dune (2021)', 'Dune', '2021'],
    ['Dune(2021)', 'Dune', '2021'],
    ['  The Matrix (1999)  ', 'The Matrix', '1999'],
  ])('reads the year on %p', (title, baseTitle, year) => {
    expect(splitMovieTitle(title)).toEqual({ baseTitle, year });
  });

  test.each(['(2026)', '  (2026)  ', 'Dune', '2001: A Space Odyssey'])(
    'keeps %p as a title',
    (title) => {
      expect(splitMovieTitle(title)).toEqual({
        baseTitle: title.trim(),
        year: null,
      });
    },
  );
});

describe('stored and displayed movie titles', () => {
  test('moves the article and keeps the year beside the title', () => {
    const stored = toStoredMovieTitle('The Matrix (1999)');
    expect(stored).toEqual({ title: 'Matrix, The', year: 1999 });
    expect(toDisplayMovieTitle('Matrix, The', 1999)).toBe('The Matrix (1999)');
    expect(toDisplayMovieTitle(stored.title, stored.year)).toBe(
      'The Matrix (1999)',
    );
  });

  test('leaves a stored title unchanged', () => {
    expect(toStoredMovieTitle('Matrix, The (1999)')).toEqual({
      title: 'Matrix, The',
      year: 1999,
    });
    expect(toStoredMovieTitle('Dune')).toEqual({ title: 'Dune', year: null });
    expect(toDisplayMovieTitle('Dune', null)).toBe('Dune');
  });

  test('shows a space before the year when the typed title omits it', () => {
    const stored = toStoredMovieTitle('Dune(2021)');
    expect(stored).toEqual({ title: 'Dune', year: 2021 });
    expect(toDisplayMovieTitle(stored.title, stored.year)).toBe('Dune (2021)');
  });
});
