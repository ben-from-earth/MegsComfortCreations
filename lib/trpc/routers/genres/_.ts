import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import { Genre as GenreEnum } from '@/lib/enums/genreEnums';
import type {
  SuccessfulGenreLinkUnlinkResponse,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
// database import
import { db as defaultDb } from '@/db/client';

// interfaces and types
import { books, genres, genresBooks } from '@/db/schema';
import { and, asc, desc, eq, isNull, sql } from 'drizzle-orm';
import { loadBookImagesById } from 'lib/media-storage/media-image-records';

const validSortKeys = ['title', 'pubYear', 'spineColor'] as const;
type SortKey = (typeof validSortKeys)[number];

function resolveSortColumn(sortKey: SortKey) {
  if (sortKey === 'title') return books.title;
  if (sortKey === 'pubYear') return books.pubYear;
  return books.spineColor;
}

export const genresRouter = router({
  getAll: adminProcedure.query(async ({ ctx }) => {
    const db = ctx.db ?? defaultDb;
    const rows = await db.select({ genre: genres.genre }).from(genres);
    const genreList = rows.map((row) => row.genre);
    return { message: 'Success', genres: genreList } as {
      message: string;
      genres: GenreEnum[];
    };
  }),

  getForBook: adminProcedure
    .input(z.object({ bookID: z.uuid() }))
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const rows = await db
        .select({ genre: genres.genre })
        .from(genres)
        .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
        .where(eq(genresBooks.bookId, input.bookID));
      const genreList = rows.map((row) => row.genre);
      return {
        message: `Successfully grabbed genres for bookID ${input.bookID}`,
        genres: genreList,
      };
    }),

  link: adminProcedure
    .input(
      z.object({
        bookID: z.uuid(),
        genres: z.array(z.string().min(1)),
      }),
    )
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const responses: SuccessfulGenreLinkUnlinkResponse[] = [];
      for (const genre of input.genres) {
        const [genreRow] = await db
          .select({ id: genres.id })
          .from(genres)
          .where(eq(genres.genre, genre));

        if (!genreRow) {
          responses.push({
            message: `Genre "${genre}" not found in database.`,
            genre,
            bookID: input.bookID,
          });
          continue;
        }

        // 2. Insert into join table
        await db.insert(genresBooks).values({
          bookId: input.bookID,
          genreId: genreRow.id,
        });
        responses.push({
          message: 'Successful genre link',
          genre,
          bookID: input.bookID,
        });
      }
      return { genreResponses: responses };
    }),

  unlink: adminProcedure
    .input(
      z.object({
        bookID: z.uuid(),
        genres: z.array(z.string().min(1)),
      }),
    )
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const responses: SuccessfulGenreLinkUnlinkResponse[] = [];
      for (const g of input.genres) {
        const [genreRow] = await db
          .select({ id: genres.id })
          .from(genres)
          .where(eq(genres.genre, g));

        if (genreRow) {
          await db
            .delete(genresBooks)
            .where(
              and(
                eq(genresBooks.bookId, input.bookID),
                eq(genresBooks.genreId, genreRow.id),
              ),
            );
        }
        responses.push({
          message: 'Successful genre unlink',
          genre: g,
          bookID: input.bookID,
        });
      }
      return { genreResponses: responses };
    }),

  paginateByGenre: adminProcedure
    .input(
      z.object({
        genre: z.string().min(1),
        limit: z.number().int().positive(),
        page: z.number().int().positive(),
        sort: z.enum(['title', 'pubYear', 'spineColor']),
        ascDesc: z.enum(['asc', 'desc']),
      }),
    )
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const offset = (input.page - 1) * input.limit;
      const sortKey = input.sort as SortKey;
      const direction = input.ascDesc.toLowerCase() === 'desc' ? 'desc' : 'asc';

      if (!validSortKeys.includes(sortKey)) {
        throw new Error(
          `Invalid sort key: ${input.sort}. Must be one of ${validSortKeys.join(', ')}`,
        );
      }

      const sortColumn = resolveSortColumn(sortKey);
      const orderByExpr =
        direction === 'desc' ? desc(sortColumn) : asc(sortColumn);

      const rows = await db
        .select({ book: books })
        .from(books)
        .innerJoin(genresBooks, eq(genresBooks.bookId, books.id))
        .innerJoin(genres, eq(genres.id, genresBooks.genreId))
        .where(eq(genres.genre, input.genre))
        .orderBy(orderByExpr)
        .limit(input.limit)
        .offset(offset);

      const [{ value: total }] = await db
        .select({ value: sql<number>`count(*)` })
        .from(books)
        .innerJoin(genresBooks, eq(genresBooks.bookId, books.id))
        .innerJoin(genres, eq(genres.id, genresBooks.genreId))
        .where(eq(genres.genre, input.genre));

      const imagesByBookId = await loadBookImagesById(
        db,
        rows.map((row) => row.book.id),
      );
      const paginatedList = rows.map((row) => ({
        ...row.book,
        images: imagesByBookId.get(row.book.id) ?? [],
      }));
      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList,
        total,
      };
      return res;
    }),

  paginateNoGenres: adminProcedure
    .input(
      z.object({
        limit: z.number().int().positive(),
        page: z.number().int().positive(),
        sort: z.enum(['title', 'pubYear', 'spineColor']),
        ascDesc: z.enum(['asc', 'desc']),
      }),
    )
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const offset = (input.page - 1) * input.limit;
      const sortKey = input.sort as SortKey;
      const direction = input.ascDesc.toLowerCase() === 'desc' ? 'desc' : 'asc';

      if (!validSortKeys.includes(sortKey)) {
        throw new Error(
          `Invalid sort key: ${input.sort}. Must be one of ${validSortKeys.join(', ')}`,
        );
      }

      const sortColumn = resolveSortColumn(sortKey);
      const orderByExpr =
        direction === 'desc' ? desc(sortColumn) : asc(sortColumn);

      const rows = await db
        .select({ book: books })
        .from(books)
        .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
        .where(isNull(genresBooks.bookId))
        .orderBy(orderByExpr)
        .limit(input.limit)
        .offset(offset);

      const [{ value: total }] = await db
        .select({ value: sql<number>`count(*)` })
        .from(books)
        .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
        .where(isNull(genresBooks.bookId));

      const imagesByBookId = await loadBookImagesById(
        db,
        rows.map((row) => row.book.id),
      );
      const paginatedList = rows.map((row) => ({
        ...row.book,
        images: imagesByBookId.get(row.book.id) ?? [],
      }));
      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList,
        total,
      };
      return res;
    }),
});
