import {
  movieTitleWithAnyYearPattern,
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
  test('moves the article and keeps the year at the end', () => {
    expect(toStoredMovieTitle('The Matrix (1999)')).toBe(
      'Matrix, The (1999)',
    );
    expect(toDisplayMovieTitle('Matrix, The (1999)')).toBe(
      'The Matrix (1999)',
    );
    expect(toDisplayMovieTitle(toStoredMovieTitle('The Matrix (1999)'))).toBe(
      'The Matrix (1999)',
    );
  });

  test('leaves a stored title unchanged', () => {
    expect(toStoredMovieTitle('Matrix, The (1999)')).toBe(
      'Matrix, The (1999)',
    );
    expect(toStoredMovieTitle('Dune')).toBe('Dune');
  });

  test('writes a space before the year when the typed title omits it', () => {
    expect(toStoredMovieTitle('Dune(2021)')).toBe('Dune (2021)');
  });
});

describe('movieTitleWithAnyYearPattern', () => {
  test('matches the stored base title plus any year, and nothing longer', () => {
    expect(movieTitleWithAnyYearPattern('Dune')).toBe(
      '^Dune \\([0-9]{4}\\)$',
    );
  });

  test('escapes regex characters in the stored base title', () => {
    expect(movieTitleWithAnyYearPattern('Dr. No')).toBe(
      '^Dr\\. No \\([0-9]{4}\\)$',
    );
    expect(movieTitleWithAnyYearPattern('(500) Days of Summer')).toBe(
      '^\\(500\\) Days of Summer \\([0-9]{4}\\)$',
    );
  });
});
