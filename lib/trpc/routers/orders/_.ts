import { TRPCError } from '@trpc/server';
import { and, asc, desc, eq, ilike, inArray, sql } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb, type Db } from '@/db/client';
import {
  albums,
  books,
  genres,
  genresBooks,
  mediaImageItems,
  movies,
  orderCovers,
  orders,
  videoGames,
} from '@/db/schema';
import { MEDIA_TYPES, type MediaType } from 'lib/constants/media-types';
import type { MovieFormat } from 'lib/constants/movie-formats';
import { toDisplayMovieTitle } from 'lib/helpers/movie-title';
import { toDisplayTitle } from 'lib/helpers/title-articles';
import { libraryTitleSearchTerm } from 'lib/trpc/routers/database/library-list-filters';
import type { MediaImageItem } from 'lib/interfaces/global-interfaces';
import {
  loadAlbumImagesById,
  loadBookImagesById,
  loadMovieImagesById,
  loadVideoGameImagesById,
} from 'lib/media-storage/media-image-records';
import { adminProcedure, router } from 'lib/trpc/trpc';

const listInput = z.object({
  name: z.string().optional(),
  number: z.string().optional(),
  title: z.string().optional(),
  type: z.enum(MEDIA_TYPES).optional(),
  min: z.number().int().nonnegative().optional(),
  max: z.number().int().nonnegative().optional(),
});

function trimmed(value: string | undefined) {
  const text = value?.trim() ?? '';
  return text === '' ? null : text;
}

function parentIdColumn(type: MediaType) {
  switch (type) {
    case 'book':
      return mediaImageItems.bookId;
    case 'movie':
      return mediaImageItems.movieId;
    case 'videoGame':
      return mediaImageItems.videoGameId;
    case 'album':
      return mediaImageItems.albumId;
  }
}

function coverCount(type: MediaType) {
  return sql`(
    select count(*)::int
    from ${orderCovers}
    inner join ${mediaImageItems}
      on ${mediaImageItems.id} = ${orderCovers.mediaImageItemId}
    where ${orderCovers.orderId} = ${orders.id}
      and ${parentIdColumn(type)} is not null
  )`;
}

function templateValue(pngFormat: number) {
  if (pngFormat === 3 || pngFormat === 5) {
    return pngFormat;
  }
  throw new TRPCError({
    code: 'INTERNAL_SERVER_ERROR',
    message: 'Stored order template is invalid.',
  });
}

function emptyCounts() {
  return { book: 0, movie: 0, videoGame: 0, album: 0 };
}

function coverType(row: {
  bookId: string | null;
  movieId: string | null;
  videoGameId: string | null;
  albumId: string | null;
}): MediaType | null {
  if (row.bookId) return 'book';
  if (row.movieId) return 'movie';
  if (row.videoGameId) return 'videoGame';
  if (row.albumId) return 'album';
  return null;
}

function uniqueIds(ids: Array<string | null>) {
  return [...new Set(ids.filter((id): id is string => id !== null))];
}

function displaySpineColor(
  images: Array<{ isDefault: boolean; spineColor: string }>,
  fallback: string,
) {
  const defaultImage = images.find((image) => image.isDefault);
  return defaultImage?.spineColor ?? images[0]?.spineColor ?? fallback;
}

type OrderCollectorBlock = {
  blockID: string;
  isDatabase: true;
  type: MediaType;
  images: ReturnType<typeof imagesForLink>;
  blockInfo: {
    title: string;
    spineColor: string;
    genres: string[];
    author?: string;
    pubYear?: number | null;
    pageCount?: number | null;
    movieFormat?: MovieFormat;
    console?: string;
    artist?: string;
    releaseDate?: string | null;
    spotifyAlbumUri?: string | null;
    codeStatus?: 'generated';
    codeGeneratedFor?: { title: string; artist: string };
  };
};

function imagesForLink(images: MediaImageItem[], coverId: string) {
  return images.flatMap((image) => {
    if (!image.id) {
      return [];
    }
    return [
      {
        id: image.id,
        url: image.url,
        selected: image.id === coverId,
        isDefault: image.isDefault,
        spineColor: image.spineColor,
        codeColor: image.codeColor ?? null,
      },
    ];
  });
}

async function loadBookGenres(db: Db, bookIds: string[]) {
  if (bookIds.length === 0) {
    return new Map<string, string[]>();
  }
  const rows = await db
    .select({ bookId: genresBooks.bookId, genre: genres.genre })
    .from(genresBooks)
    .innerJoin(genres, eq(genresBooks.genreId, genres.id))
    .where(inArray(genresBooks.bookId, bookIds));
  const genresByBookId = new Map<string, string[]>();
  for (const row of rows) {
    const current = genresByBookId.get(row.bookId) ?? [];
    current.push(row.genre);
    genresByBookId.set(row.bookId, current);
  }
  return genresByBookId;
}

export const ordersRouter = router({
  list: adminProcedure.input(listInput).query(async ({ input, ctx }) => {
    const db = ctx.db ?? defaultDb;
    const [{ total }] = await db
      .select({ total: sql<number>`count(*)::int` })
      .from(orders);
    const min = input.min;
    const max = input.max;
    if (min !== undefined && max !== undefined && min > max) {
      return { total, rows: [] };
    }

    const conditions = [];
    const name = trimmed(input.name);
    if (name) {
      conditions.push(ilike(orders.customerName, `%${name}%`));
    }
    const number = trimmed(input.number);
    if (number) {
      conditions.push(ilike(orders.orderNumber, `%${number}%`));
    }
    const title = trimmed(input.title);
    if (title) {
      const pattern = `%${libraryTitleSearchTerm(title)}%`;
      conditions.push(
        sql`exists (
          select 1
          from ${orderCovers}
          inner join ${mediaImageItems}
            on ${mediaImageItems.id} = ${orderCovers.mediaImageItemId}
          left join ${books} on ${books.id} = ${mediaImageItems.bookId}
          left join ${movies} on ${movies.id} = ${mediaImageItems.movieId}
          left join ${videoGames}
            on ${videoGames.id} = ${mediaImageItems.videoGameId}
          left join ${albums} on ${albums.id} = ${mediaImageItems.albumId}
          where ${orderCovers.orderId} = ${orders.id}
            and (
              ${books.title} ilike ${pattern}
              or ${movies.title} ilike ${pattern}
              or ${videoGames.title} ilike ${pattern}
              or ${albums.title} ilike ${pattern}
            )
        )`,
      );
    }
    if (input.type && (min !== undefined || max !== undefined)) {
      const count = coverCount(input.type);
      if (min !== undefined) {
        conditions.push(sql`${count} >= ${min}`);
      }
      if (max !== undefined) {
        conditions.push(sql`${count} <= ${max}`);
      }
    }

    const rows = await db
      .select({
        id: orders.id,
        customerName: orders.customerName,
        orderNumber: orders.orderNumber,
        pngFormat: orders.pngFormat,
        repeatCount: orders.repeatCount,
        fileUrl: orders.fileUrl,
        fileName: orders.fileName,
        updatedAt: orders.updatedAt,
      })
      .from(orders)
      .where(conditions.length > 0 ? and(...conditions) : undefined)
      .orderBy(desc(orders.updatedAt));

    if (rows.length === 0) {
      return { total, rows: [] };
    }

    const covers = await db
      .select({
        orderId: orderCovers.orderId,
        position: orderCovers.position,
        url: mediaImageItems.path,
        bookId: mediaImageItems.bookId,
        movieId: mediaImageItems.movieId,
        videoGameId: mediaImageItems.videoGameId,
        albumId: mediaImageItems.albumId,
      })
      .from(orderCovers)
      .innerJoin(
        mediaImageItems,
        eq(mediaImageItems.id, orderCovers.mediaImageItemId),
      )
      .where(
        inArray(
          orderCovers.orderId,
          rows.map((row) => row.id),
        ),
      )
      .orderBy(asc(orderCovers.orderId), asc(orderCovers.position));

    const summaryByOrderId = new Map<
      string,
      {
        counts: ReturnType<typeof emptyCounts>;
        firstCover: { url: string } | null;
      }
    >();
    for (const cover of covers) {
      const summary = summaryByOrderId.get(cover.orderId) ?? {
        counts: emptyCounts(),
        firstCover: null,
      };
      const type = coverType(cover);
      if (type) {
        summary.counts[type] += 1;
      }
      if (!summary.firstCover) {
        summary.firstCover = { url: cover.url };
      }
      summaryByOrderId.set(cover.orderId, summary);
    }

    return {
      total,
      rows: rows.map((row) => {
        const summary = summaryByOrderId.get(row.id);
        return {
          id: row.id,
          customerName: row.customerName,
          orderNumber: row.orderNumber,
          template: templateValue(row.pngFormat),
          repeatCount: row.repeatCount,
          fileUrl: row.fileUrl,
          fileName: row.fileName,
          updatedAt: row.updatedAt,
          counts: summary?.counts ?? emptyCounts(),
          firstCover: summary?.firstCover ?? null,
        };
      }),
    };
  }),

  getForCollector: adminProcedure
    .input(z.object({ id: z.uuid() }))
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const [order] = await db
        .select({
          id: orders.id,
          customerName: orders.customerName,
          orderNumber: orders.orderNumber,
          pngFormat: orders.pngFormat,
          repeatCount: orders.repeatCount,
        })
        .from(orders)
        .where(eq(orders.id, input.id))
        .limit(1);
      if (!order) {
        throw new TRPCError({
          code: 'NOT_FOUND',
          message: 'This order no longer exists.',
        });
      }

      const links = await db
        .select({
          coverId: orderCovers.mediaImageItemId,
          bookId: mediaImageItems.bookId,
          movieId: mediaImageItems.movieId,
          videoGameId: mediaImageItems.videoGameId,
          albumId: mediaImageItems.albumId,
        })
        .from(orderCovers)
        .innerJoin(
          mediaImageItems,
          eq(mediaImageItems.id, orderCovers.mediaImageItemId),
        )
        .where(eq(orderCovers.orderId, order.id))
        .orderBy(asc(orderCovers.position));

      const bookIds = uniqueIds(links.map((link) => link.bookId));
      const movieIds = uniqueIds(links.map((link) => link.movieId));
      const videoGameIds = uniqueIds(links.map((link) => link.videoGameId));
      const albumIds = uniqueIds(links.map((link) => link.albumId));

      const [
        bookRows,
        movieRows,
        videoGameRows,
        albumRows,
        bookImages,
        movieImages,
        videoGameImages,
        albumImages,
        genresByBookId,
      ] = await Promise.all([
        bookIds.length === 0
          ? Promise.resolve([])
          : db.select().from(books).where(inArray(books.id, bookIds)),
        movieIds.length === 0
          ? Promise.resolve([])
          : db.select().from(movies).where(inArray(movies.id, movieIds)),
        videoGameIds.length === 0
          ? Promise.resolve([])
          : db
              .select()
              .from(videoGames)
              .where(inArray(videoGames.id, videoGameIds)),
        albumIds.length === 0
          ? Promise.resolve([])
          : db.select().from(albums).where(inArray(albums.id, albumIds)),
        loadBookImagesById(db, bookIds),
        loadMovieImagesById(db, movieIds),
        loadVideoGameImagesById(db, videoGameIds),
        loadAlbumImagesById(db, albumIds),
        loadBookGenres(db, bookIds),
      ]);

      const bookById = new Map(bookRows.map((row) => [row.id, row]));
      const movieById = new Map(movieRows.map((row) => [row.id, row]));
      const videoGameById = new Map(videoGameRows.map((row) => [row.id, row]));
      const albumById = new Map(albumRows.map((row) => [row.id, row]));

      const loadedLibraryIds = new Set<string>();
      const blocks = links.flatMap((link): OrderCollectorBlock[] => {
        const libraryId =
          link.bookId ?? link.movieId ?? link.videoGameId ?? link.albumId;
        if (!libraryId || loadedLibraryIds.has(libraryId)) {
          return [];
        }
        loadedLibraryIds.add(libraryId);
        if (link.bookId) {
          const book = bookById.get(link.bookId);
          if (!book) return [];
          const images = imagesForLink(
            bookImages.get(book.id) ?? [],
            link.coverId,
          );
          return [
            {
              blockID: book.id,
              isDatabase: true as const,
              type: 'book' as const,
              images,
              blockInfo: {
                title: toDisplayTitle(book.title),
                spineColor: displaySpineColor(images, book.spineColor),
                genres: genresByBookId.get(book.id) ?? [],
                author: book.author,
                pubYear: book.pubYear,
                pageCount: book.pageCount,
              },
            },
          ];
        }
        if (link.movieId) {
          const movie = movieById.get(link.movieId);
          if (!movie) return [];
          const images = imagesForLink(
            movieImages.get(movie.id) ?? [],
            link.coverId,
          );
          return [
            {
              blockID: movie.id,
              isDatabase: true as const,
              type: 'movie' as const,
              images,
              blockInfo: {
                title: toDisplayMovieTitle(movie.title, movie.year),
                spineColor: displaySpineColor(images, movie.spineColor),
                genres: [],
                movieFormat: movie.movieFormat,
              },
            },
          ];
        }
        if (link.videoGameId) {
          const videoGame = videoGameById.get(link.videoGameId);
          if (!videoGame) return [];
          const images = imagesForLink(
            videoGameImages.get(videoGame.id) ?? [],
            link.coverId,
          );
          return [
            {
              blockID: videoGame.id,
              isDatabase: true as const,
              type: 'videoGame' as const,
              images,
              blockInfo: {
                title: toDisplayMovieTitle(videoGame.title, videoGame.year),
                spineColor: displaySpineColor(images, videoGame.spineColor),
                genres: [],
                console: videoGame.console,
              },
            },
          ];
        }
        if (link.albumId) {
          const album = albumById.get(link.albumId);
          if (!album) return [];
          const images = imagesForLink(
            albumImages.get(album.id) ?? [],
            link.coverId,
          );
          const title = toDisplayTitle(album.title);
          return [
            {
              blockID: album.id,
              isDatabase: true as const,
              type: 'album' as const,
              images,
              blockInfo: {
                title,
                spineColor: displaySpineColor(images, album.spineColor),
                genres: [],
                artist: album.artist,
                releaseDate: album.releaseDate,
                spotifyAlbumUri: album.spotifyAlbumUri,
                ...(album.spotifyAlbumUri
                  ? {
                      codeStatus: 'generated' as const,
                      codeGeneratedFor: { title, artist: album.artist },
                    }
                  : {}),
              },
            },
          ];
        }
        return [];
      });
      const blockIds = new Set(blocks.map((block) => block.blockID));

      return {
        customerName: order.customerName,
        orderNumber: order.orderNumber,
        template: templateValue(order.pngFormat),
        repeatCount: order.repeatCount,
        blocks,
        pngExportBlockIds: links.flatMap((link) => {
          const libraryId =
            link.bookId ?? link.movieId ?? link.videoGameId ?? link.albumId;
          return libraryId && blockIds.has(libraryId) ? [libraryId] : [];
        }),
      };
    }),
});
