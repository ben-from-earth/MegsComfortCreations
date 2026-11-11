import { and, eq, ilike, or, sql } from 'drizzle-orm';
import { Db } from '@/db/client';
import { MediaType } from 'lib/constants/media-types';
import { type MovieFormat } from 'lib/constants/movie-formats';
import {
  movieTitleWithAnyYearPattern,
  splitMovieTitle,
  toDisplayMovieTitle,
  toStoredMovieTitle,
} from 'lib/helpers/movie-title';
import { toDisplayTitle, toStoredTitle } from 'lib/helpers/title-articles';
import { books, otherMedia } from '@/db/schema';
import {
  loadBookImagesById,
  loadOtherMediaImagesById,
} from 'lib/media-storage/media-image-records';

function otherMediaTitleCondition(type: MediaType, title: string) {
  if (type !== 'movie') {
    return ilike(otherMedia.title, toStoredTitle(title));
  }

  const storedTitle = toStoredMovieTitle(title);
  if (splitMovieTitle(title).year) {
    return ilike(otherMedia.title, storedTitle);
  }

  return or(
    ilike(otherMedia.title, storedTitle),
    sql`${otherMedia.title} ~* ${movieTitleWithAnyYearPattern(storedTitle)}`,
  );
}

export function resolveBookSearchFilters(title: string, author?: string) {
  const rearrangedTitle = toStoredTitle(title);
  const normalizedAuthor = author?.trim() || undefined;
  return {
    rearrangedTitle,
    author: normalizedAuthor,
  };
}

export async function searchByTitle(
  db: Db,
  type: MediaType,
  title: string,
  author?: string,
  movieFormat?: MovieFormat,
) {
  const { rearrangedTitle, author: normalizedAuthor } = resolveBookSearchFilters(
    title,
    author,
  );
  const result =
    type === 'book'
      ? await (normalizedAuthor
          ? db
              .select()
              .from(books)
              .where(
                and(
                  ilike(books.title, rearrangedTitle),
                  ilike(books.author, normalizedAuthor),
                ),
              )
          : db
              .select()
              .from(books)
              .where(ilike(books.title, rearrangedTitle)))
      : await db
          .select()
          .from(otherMedia)
          .where(
            and(
              eq(otherMedia.mediaType, type),
              otherMediaTitleCondition(type, title),
              type === 'movie'
                ? eq(otherMedia.movieFormat, movieFormat ?? 'dvd')
                : undefined,
            ),
          );
  const normalizedResult =
    type === 'book'
      ? await (async () => {
          const imagesByBookId = await loadBookImagesById(
            db,
            result.map((row) => row.id),
          );
          return result.map((row) => ({
            ...row,
            title: toDisplayTitle(row.title),
            images: imagesByBookId.get(row.id) ?? [],
          }));
        })()
      : await (async () => {
          const imagesByOtherMediaId = await loadOtherMediaImagesById(
            db,
            result.map((row) => row.id),
          );
          return result.map((row) => ({
            ...row,
            title:
              type === 'movie'
                ? toDisplayMovieTitle(row.title)
                : toDisplayTitle(row.title),
            images: imagesByOtherMediaId.get(row.id) ?? [],
          }));
        })();
  return {
    foundMediaList: normalizedResult,
    total: normalizedResult.length,
  };
}
