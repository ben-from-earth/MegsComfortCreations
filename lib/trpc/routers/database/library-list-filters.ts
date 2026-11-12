import { MediaType } from 'lib/constants/media-types';
import { splitMovieTitle } from 'lib/helpers/movie-title';
import { toTitleSearchTerm } from 'lib/helpers/title-articles';

export function libraryTitleSearchTerm(title: string) {
  return toTitleSearchTerm(splitMovieTitle(title).baseTitle);
}

export type LibraryYearConstraint =
  | { kind: 'any' }
  | { kind: 'missing' }
  | { kind: 'year'; year: number };

export function libraryYearConstraint(
  year: number | undefined,
  missingYear: boolean | undefined,
): LibraryYearConstraint {
  if (missingYear === true) {
    return { kind: 'missing' };
  }
  if (year === undefined) {
    return { kind: 'any' };
  }
  return { kind: 'year', year };
}

export function libraryPaginationFilterError(input: {
  type: MediaType;
  year?: number;
  missingYear?: boolean;
  artist?: string;
  releaseDateFrom?: string;
  releaseDateTo?: string;
}): string | null {
  const yearFilter = input.year !== undefined || input.missingYear === true;
  if (yearFilter && input.type !== 'movie' && input.type !== 'videoGame') {
    return 'Year filter is only supported for movies and video games';
  }

  const artist = input.artist?.trim() ?? '';
  const releaseFilter =
    input.releaseDateFrom !== undefined || input.releaseDateTo !== undefined;
  if ((artist !== '' || releaseFilter) && input.type !== 'album') {
    return 'Artist and release date filters are only supported for albums';
  }

  return null;
}
