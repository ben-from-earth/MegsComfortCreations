import { and, eq, ilike, type Column } from 'drizzle-orm';
import { Db } from '@/db/client';
import { MediaType } from 'lib/constants/media-types';
import { type MovieFormat } from 'lib/constants/movie-formats';
import {
  toDisplayMovieTitle,
  toStoredMovieTitle,
} from 'lib/helpers/movie-title';
import { toDisplayTitle, toStoredTitle } from 'lib/helpers/title-articles';
import { albums, books, movies, videoGames } from '@/db/schema';
import type { MediaImageItem, PostSavedMediaItem } from 'lib/interfaces/global-interfaces';
import {
  loadAlbumImagesById,
  loadBookImagesById,
  loadMovieImagesById,
  loadVideoGameImagesById,
} from 'lib/media-storage/media-image-records';

function titledMediaSearchCondition(
  titleColumn: Column,
  yearColumn: Column,
  title: string,
) {
  const stored = toStoredMovieTitle(title);
  const titleMatch = ilike(titleColumn, stored.title);
  if (stored.year === null) {
    return titleMatch;
  }
  return and(titleMatch, eq(yearColumn, stored.year));
}

export function resolveBookSearchFilters(title: string, author?: string) {
  const rearrangedTitle = toStoredTitle(title);
  const normalizedAuthor = author?.trim() || undefined;
  return {
    rearrangedTitle,
    author: normalizedAuthor,
  };
}

async function loadImages<T extends { id: string }>(
  db: Db,
  rows: T[],
  load: (db: Db, ids: string[]) => Promise<Map<string, MediaImageItem[]>>,
) {
  return load(
    db,
    rows.map((row) => row.id),
  );
}

export async function searchByTitle(
  db: Db,
  type: MediaType,
  title: string,
  author?: string,
  movieFormat?: MovieFormat,
  spotifyAlbumUri?: string,
): Promise<{ foundMediaList: PostSavedMediaItem[]; total: number }> {
  const { rearrangedTitle, author: normalizedAuthor } = resolveBookSearchFilters(
    title,
    author,
  );
  const spotifyUri = spotifyAlbumUri?.trim() || undefined;

  if (type === 'book') {
    const rows = await db
      .select()
      .from(books)
      .where(
        normalizedAuthor
          ? and(
              ilike(books.title, rearrangedTitle),
              ilike(books.author, normalizedAuthor),
            )
          : ilike(books.title, rearrangedTitle),
      );
    const imagesById = await loadImages(db, rows, loadBookImagesById);
    const foundMediaList: PostSavedMediaItem[] = rows.map((row) => ({
      mediaType: 'book',
      id: row.id,
      title: toDisplayTitle(row.title),
      spineColor: row.spineColor,
      images: imagesById.get(row.id) ?? [],
      author: row.author,
      pageCount: row.pageCount,
      pubYear: row.pubYear,
    }));
    return { foundMediaList, total: foundMediaList.length };
  }

  if (type === 'movie') {
    const rows = await db
      .select()
      .from(movies)
      .where(
        and(
          titledMediaSearchCondition(movies.title, movies.year, title),
          eq(movies.movieFormat, movieFormat ?? 'dvd'),
        ),
      );
    const imagesById = await loadImages(db, rows, loadMovieImagesById);
    const foundMediaList: PostSavedMediaItem[] = rows.map((row) => ({
      mediaType: 'movie',
      id: row.id,
      title: toDisplayMovieTitle(row.title, row.year),
      spineColor: row.spineColor,
      images: imagesById.get(row.id) ?? [],
      movieFormat: row.movieFormat,
    }));
    return { foundMediaList, total: foundMediaList.length };
  }

  if (type === 'videoGame') {
    const rows = await db
      .select()
      .from(videoGames)
      .where(titledMediaSearchCondition(videoGames.title, videoGames.year, title));
    const imagesById = await loadImages(db, rows, loadVideoGameImagesById);
    const foundMediaList: PostSavedMediaItem[] = rows.map((row) => ({
      mediaType: 'videoGame',
      id: row.id,
      title: toDisplayMovieTitle(row.title, row.year),
      spineColor: row.spineColor,
      images: imagesById.get(row.id) ?? [],
    }));
    return { foundMediaList, total: foundMediaList.length };
  }

  if (spotifyUri) {
    const rows = await db
      .select()
      .from(albums)
      .where(eq(albums.spotifyAlbumUri, spotifyUri));
    const imagesById = await loadImages(db, rows, loadAlbumImagesById);
    const foundMediaList: PostSavedMediaItem[] = rows.map((row) => ({
      mediaType: 'album',
      id: row.id,
      title: toDisplayTitle(row.title),
      spineColor: row.spineColor,
      images: imagesById.get(row.id) ?? [],
      artist: row.artist,
      releaseDate: row.releaseDate,
      spotifyAlbumUri: row.spotifyAlbumUri,
    }));
    return { foundMediaList, total: foundMediaList.length };
  }

  const rows = await db
    .select()
    .from(albums)
    .where(
      and(
        ilike(albums.title, rearrangedTitle),
        normalizedAuthor ? ilike(albums.artist, normalizedAuthor) : undefined,
      ),
    );
  const imagesById = await loadImages(db, rows, loadAlbumImagesById);
  const foundMediaList: PostSavedMediaItem[] = rows.map((row) => ({
    mediaType: 'album',
    id: row.id,
    title: toDisplayTitle(row.title),
    spineColor: row.spineColor,
    images: imagesById.get(row.id) ?? [],
    artist: row.artist,
    releaseDate: row.releaseDate,
    spotifyAlbumUri: row.spotifyAlbumUri,
  }));
  return { foundMediaList, total: foundMediaList.length };
}
