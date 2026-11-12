import {
  libraryPaginationFilterError,
  libraryTitleSearchTerm,
  libraryYearConstraint,
} from 'lib/trpc/routers/database/library-list-filters';

describe('libraryTitleSearchTerm', () => {
  test('drops a trailing year and a leading article', () => {
    expect(libraryTitleSearchTerm('The Matrix (1999)')).toBe('Matrix');
    expect(libraryTitleSearchTerm('Dune (2021)')).toBe('Dune');
    expect(libraryTitleSearchTerm('Dune(2021)')).toBe('Dune');
  });

  test('keeps a title that has no year suffix', () => {
    expect(libraryTitleSearchTerm('Dune')).toBe('Dune');
    expect(libraryTitleSearchTerm('(2026)')).toBe('(2026)');
    expect(libraryTitleSearchTerm('2001: A Space Odyssey')).toBe(
      '2001: A Space Odyssey',
    );
  });
});

describe('libraryYearConstraint', () => {
  test('leaves the year open when neither control is set', () => {
    expect(libraryYearConstraint(undefined, undefined)).toEqual({
      kind: 'any',
    });
    expect(libraryYearConstraint(undefined, false)).toEqual({ kind: 'any' });
  });

  test('matches one year', () => {
    expect(libraryYearConstraint(1999, undefined)).toEqual({
      kind: 'year',
      year: 1999,
    });
    expect(libraryYearConstraint(1999, false)).toEqual({
      kind: 'year',
      year: 1999,
    });
  });

  test('a request for no year wins over a year value', () => {
    expect(libraryYearConstraint(1999, true)).toEqual({ kind: 'missing' });
    expect(libraryYearConstraint(undefined, true)).toEqual({ kind: 'missing' });
  });
});

describe('libraryPaginationFilterError', () => {
  test('allows each type its own filters', () => {
    expect(
      libraryPaginationFilterError({ type: 'book', missingYear: false }),
    ).toBeNull();
    expect(
      libraryPaginationFilterError({ type: 'movie', year: 1999 }),
    ).toBeNull();
    expect(
      libraryPaginationFilterError({ type: 'videoGame', missingYear: true }),
    ).toBeNull();
    expect(
      libraryPaginationFilterError({
        type: 'album',
        artist: 'Fleetwood Mac',
        releaseDateFrom: '1977-01-01',
        releaseDateTo: '1977-12-31',
      }),
    ).toBeNull();
  });

  test('rejects a year filter on books and albums', () => {
    expect(
      libraryPaginationFilterError({ type: 'book', year: 1999 }),
    ).toBe('Year filter is only supported for movies and video games');
    expect(
      libraryPaginationFilterError({ type: 'album', missingYear: true }),
    ).toBe('Year filter is only supported for movies and video games');
  });

  test('rejects an artist or release range outside albums', () => {
    const message =
      'Artist and release date filters are only supported for albums';
    expect(
      libraryPaginationFilterError({ type: 'movie', artist: 'Prince' }),
    ).toBe(message);
    expect(
      libraryPaginationFilterError({
        type: 'book',
        releaseDateFrom: '1977-02-04',
      }),
    ).toBe(message);
    expect(
      libraryPaginationFilterError({
        type: 'videoGame',
        releaseDateTo: '1977-02-04',
      }),
    ).toBe(message);
  });

  test('ignores a blank artist', () => {
    expect(
      libraryPaginationFilterError({ type: 'movie', artist: '   ' }),
    ).toBeNull();
  });
});
