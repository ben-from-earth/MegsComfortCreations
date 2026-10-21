export const allGenres = [
  "Children's Fiction",
  'Middle-Grade Fiction',
  'Young Adult Fiction',
  'New Adult Fiction',
  'Romance',
  'Contemporary Fiction',
  'Spicy Romance (18+)',
  'LGBTQ',
  'Romantasy',
  'Fantasy',
  'Historical Fiction',
  'Mystery',
  'Thriller',
  'Horror',
  'Science Fiction',
  'Classic Literature',
  'Memoir',
  'History',
  'Philosophy',
  'Anthology',
] as const;

export type Genre = (typeof allGenres)[number];
export const NO_GENRE_FILTER = 'none' as const;
