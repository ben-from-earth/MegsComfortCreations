export const MOVIE_FORMATS = ['dvd', 'vhs', 'bluRay'] as const;

export type MovieFormat = (typeof MOVIE_FORMATS)[number];

export const MOVIE_FORMAT_OPTIONS = [
  { value: 'dvd', label: 'DVD', searchToken: 'dvd' },
  { value: 'vhs', label: 'VHS', searchToken: 'vhs' },
  { value: 'bluRay', label: 'Blu-Ray', searchToken: 'blu-ray' },
] as const satisfies ReadonlyArray<{
  value: MovieFormat;
  label: string;
  searchToken: string;
}>;
