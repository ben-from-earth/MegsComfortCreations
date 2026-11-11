export const MOVIE_FORMATS = ['dvd', 'vhs'] as const;

export type MovieFormat = (typeof MOVIE_FORMATS)[number];

export const MOVIE_FORMAT_OPTIONS = [
  { value: 'dvd', label: 'DVD', searchToken: 'dvd' },
  { value: 'vhs', label: 'VHS', searchToken: 'vhs' },
] as const satisfies ReadonlyArray<{
  value: MovieFormat;
  label: string;
  searchToken: string;
}>;
