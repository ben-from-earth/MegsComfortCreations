import {
  toDisplayTitle,
  toStoredTitle,
} from 'lib/helpers/title-articles';

const YEAR_SUFFIX = /^\((\d{4})\)$/;

type SplitMovieTitle = {
  baseTitle: string;
  year: string | null;
};

export function splitMovieTitle(title: string): SplitMovieTitle {
  const trimmed = title.trim();
  if (trimmed.length < 7) {
    return { baseTitle: trimmed, year: null };
  }

  const suffix = trimmed.slice(-6);
  const match = YEAR_SUFFIX.exec(suffix);
  if (!match) {
    return { baseTitle: trimmed, year: null };
  }

  const baseTitle = trimmed.slice(0, -6).trim();
  if (baseTitle.length === 0) {
    return { baseTitle: trimmed, year: null };
  }

  return { baseTitle, year: match[1] };
}

export function toStoredMovieTitle(title: string): {
  title: string;
  year: number | null;
} {
  const { baseTitle, year } = splitMovieTitle(title);
  return {
    title: toStoredTitle(baseTitle),
    year: year === null ? null : Number(year),
  };
}

export function toDisplayMovieTitle(
  title: string,
  year: number | null,
): string {
  const displayTitle = toDisplayTitle(title);
  return year === null ? displayTitle : `${displayTitle} (${year})`;
}
