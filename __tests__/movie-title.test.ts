import { toTitleSearchTerm } from 'lib/helpers/title-articles';
import {
  movieTitlesMatch,
  movieYearlessTitlePattern,
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

describe('movieTitlesMatch', () => {
  test('a title without a year matches every stored year', () => {
    expect(movieTitlesMatch('Dune', 'Dune')).toBe(true);
    expect(movieTitlesMatch('Dune (1984)', 'Dune')).toBe(true);
    expect(movieTitlesMatch('Dune (2021)', 'Dune')).toBe(true);
    expect(movieTitlesMatch('Matrix, The', 'The Matrix')).toBe(true);
    expect(movieTitlesMatch('Matrix, The (1999)', 'The Matrix')).toBe(true);
  });

  test('a title with a year matches only that year', () => {
    expect(movieTitlesMatch('Dune (2021)', 'Dune (2021)')).toBe(true);
    expect(movieTitlesMatch('Dune (2021)', 'Dune(2021)')).toBe(true);
    expect(movieTitlesMatch('Matrix, The (1999)', 'The Matrix (1999)')).toBe(
      true,
    );
    expect(movieTitlesMatch('Dune', 'Dune (2021)')).toBe(false);
    expect(movieTitlesMatch('Dune (1984)', 'Dune (2021)')).toBe(false);
    expect(movieTitlesMatch('Matrix, The (2003)', 'The Matrix (1999)')).toBe(
      false,
    );
  });

  test('does not match a longer title that only starts with the query', () => {
    expect(movieTitlesMatch('Dune: Part Two (2024)', 'Dune')).toBe(false);
  });
});

describe('movie search text', () => {
  test('library search drops the year before the article search term', () => {
    const { baseTitle, year } = splitMovieTitle('The Matrix (1999)');
    expect(toTitleSearchTerm(baseTitle)).toBe('Matrix');
    expect(year).toBe('1999');
  });

  test('year-less collector search escapes regex characters in the base title', () => {
    expect(movieYearlessTitlePattern('Dr. No')).toBe(
      '^Dr\\. No \\([0-9]{4}\\)$',
    );
    expect(movieYearlessTitlePattern('(500) Days of Summer')).toBe(
      '^\\(500\\) Days of Summer \\([0-9]{4}\\)$',
    );
  });
});
