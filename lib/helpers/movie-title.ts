import {
  toDisplayTitle,
  toStoredTitle,
} from 'lib/helpers/title-articles';

const YEAR_SUFFIX = /^\((\d{4})\)$/;
const REGEX_METACHARACTERS = /[\\.^$*+?()[\]{}|]/g;

export type SplitMovieTitle = {
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
  if (!match?.[1]) {
    return { baseTitle: trimmed, year: null };
  }

  const baseTitle = trimmed.slice(0, -6).trim();
  if (baseTitle.length === 0) {
    return { baseTitle: trimmed, year: null };
  }

  return { baseTitle, year: match[1] };
}

function withYear(baseTitle: string, year: string | null) {
  return year ? `${baseTitle} (${year})` : baseTitle;
}

export function toStoredMovieTitle(title: string): string {
  const { baseTitle, year } = splitMovieTitle(title);
  return withYear(toStoredTitle(baseTitle), year);
}

export function toDisplayMovieTitle(title: string): string {
  const { baseTitle, year } = splitMovieTitle(title);
  return withYear(toDisplayTitle(baseTitle), year);
}

export function movieYearlessTitlePattern(storedBaseTitle: string): string {
  const escapedBase = storedBaseTitle.replace(REGEX_METACHARACTERS, '\\$&');
  return `^${escapedBase} \\([0-9]{4}\\)$`;
}

export function movieTitlesMatch(storedTitle: string, query: string): boolean {
  const stored = splitMovieTitle(storedTitle);
  const queried = splitMovieTitle(query);
  const storedBase = toStoredTitle(stored.baseTitle).toLowerCase();
  const queryBase = toStoredTitle(queried.baseTitle).toLowerCase();
  if (storedBase !== queryBase) {
    return false;
  }
  if (queried.year === null) {
    return true;
  }
  return stored.year === queried.year;
}
