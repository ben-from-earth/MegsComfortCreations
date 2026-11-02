const VISIBLE_GENRE_COUNT = 2;

export function splitVisibleGenres(genres: string[]) {
  return {
    visibleGenres: genres.slice(0, VISIBLE_GENRE_COUNT),
    moreCount: Math.max(genres.length - VISIBLE_GENRE_COUNT, 0),
  };
}

export function filterGenresBySearch(
  allGenres: readonly string[],
  search: string,
) {
  const query = search.trim().toLowerCase();
  return allGenres.filter((genre) => genre.toLowerCase().includes(query));
}
