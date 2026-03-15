export const OTHER_MEDIA_TYPES = ['movie', 'videoGame', 'album'] as const;
export type OtherMediaType = (typeof OTHER_MEDIA_TYPES)[number];

export const MEDIA_TYPES = ['book', ...OTHER_MEDIA_TYPES] as const;
export type MediaType = (typeof MEDIA_TYPES)[number];

export const MEDIA_TYPE_DEFINITIONS = [
  { mediaType: 'book', label: 'Book', plural: 'Books' },
  { mediaType: 'movie', label: 'Movie', plural: 'Movies' },
  { mediaType: 'videoGame', label: 'Video Game', plural: 'Video Games' },
  { mediaType: 'album', label: 'Album', plural: 'Albums' },
] as const;
