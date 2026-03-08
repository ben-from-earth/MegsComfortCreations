import { adminProcedure, router } from 'lib/trpc/trpc';
import { and, asc, desc, ilike, isNull, sql, eq } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import {
  albums,
  books,
  genres,
  genresBooks,
  googleApiQueryUsage,
  movies,
  videoGames,
} from '@/db/schema';
import type {
  DatabaseSaveServerResponse,
  SuccessfulPaginationResponse,
  BookRow,
} from 'lib/interfaces/globalInterfaces';
import { validate } from 'jsonschema';
import bookCreateSchema from 'lib/database/schemas/bookCreateSchema.json';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { searchByTitle } from './actions/search-by-title';
import { collectedBlockInformationSchema } from '@/mediacollector/collector-form/collectorFormSchema';
import { allGenres } from '@/lib/enums/genreEnums';

const mediaType = z.enum(['book', 'movie', 'videoGame', 'album']);

const tableMap = {
  book: books,
  movie: movies,
  videoGame: videoGames,
  album: albums,
} as const;

export const databaseRouter = router({
  searchByTitle: adminProcedure
    .input(z.object({ type: mediaType, title: z.string().min(1) }))
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const { type, title } = input;

      return await searchByTitle(db, type, title);
    }),
  getPaginated: adminProcedure
    .input(
      z.object({
        type: mediaType,
        title: z.string().optional(),
        limit: z.number().int().positive(),
        page: z.number().int().positive(),
        sort: z.enum(['title', 'author', 'pubYear', 'spineColor']),
        ascDesc: z.enum(['asc', 'desc']),
        genre: z.enum([...allGenres, '', 'None']),
      }),
    )
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const { limit, page, sort, ascDesc, genre, title } = input;
      // const table = tableMap[type];
      const offset = (page - 1) * limit;

      // Determine the correct sort column per media type
      const sortColumn = (() => {
        switch (sort) {
          case 'author':
            return books.author;
          case 'pubYear':
            return books.pubYear;
          case 'spineColor':
            return books.spineColor;
          default:
            return books.title;
        }
      })();

      const orderExpr = ascDesc === 'asc' ? asc(sortColumn) : desc(sortColumn);

      const genreFilter = (() => {
        if (genre === '') {
          // no filter
          return undefined;
        }

        if (genre === 'None') {
          // only books with no row in the link table
          return isNull(genresBooks.bookId); // or isNull(genresBooks.genreId)
        }

        // specific genre
        return eq(genres.genre, genre);
      })();

      const rows = await db
        .select({
          id: books.id,
          title: books.title,
          author: books.author,
          pageCount: books.pageCount,
          pubYear: books.pubYear,
          spineColor: books.spineColor,
          imageUrls: books.imageUrls,
          genre: genres.genre, // will be null when no link
        })
        .from(books)
        .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
        .leftJoin(genres, eq(genresBooks.genreId, genres.id))
        .where(and(genreFilter, ilike(books.title, `%${title}%`)))
        .orderBy(orderExpr)
        .limit(limit)
        .offset(offset);

      const [{ count }] = await db
        .select({ count: sql<number>`cast(count(*) as int)` })
        .from(books);

      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList: rows,
        total: count,
      };
      return res;
    }),

  deleteByTitle: adminProcedure
    .input(z.object({ type: mediaType, title: z.string().min(1) }))
    .mutation(async ({ input, ctx }) => {
      const { type, title } = input;
      const table = tableMap[type];
      const db = ctx.db ?? defaultDb;

      const deleted = await db
        .delete(table)
        .where(ilike(table.title, title))
        .returning({ id: table.id });

      if (deleted.length === 0) {
        return {
          message: `No item with title: ${title} in the ${type} database exists`,
        };
      }
      return { message: `Successfully deleted ${title}` };
    }),

  getQueryCount: adminProcedure
    .input(z.object({ date: z.string().min(1) }))
    .query(async ({ input, ctx }) => {
      const { date } = input;
      const db = ctx.db ?? defaultDb;

      const [row] = await db
        .select()
        .from(googleApiQueryUsage)
        .where(eq(googleApiQueryUsage.date, date))
        .limit(1);

      if (!row) {
        return { date, queryCount: 0 };
      }
      return { date: row.date, queryCount: row.queryCount };
    }),

  save: adminProcedure
    .input(z.array(collectedBlockInformationSchema))
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;

      const results: DatabaseSaveServerResponse = [];

      for (const book of input) {
        if (book.isDatabase) {
          continue;
        }
        book.images = book.images.filter((img) => img.selected);
        const validatedData = collectedBlockInformationSchema.safeParse(book);
        if (!validatedData.success) {
          const tree = z.treeifyError(validatedData.error);
          results.push({
            error: 'Schema Violation',
            message: 'Schema violation(s) during save request',
            errors: tree.errors,
            title: book.blockInfo.title,
          });
        } else {
          const data = validatedData.data;

          const insertData = {
            title: titleRearrange(data.blockInfo.title),
            author: data.blockInfo.author!,
            pageCount: data.blockInfo.pageCount ?? null,
            pubYear: data.blockInfo.pubYear ?? null,
            spineColor: data.blockInfo.spineColor,
            imageUrls: data.images.map((img) => img.url),
          };

          const [book] = await db.insert(books).values(insertData).returning();

          if (!book) {
            results.push({
              title: data.blockInfo.title,
              error: 'Database Insertion Error',
              message: 'An error occurred while trying to save to the database',
              errors: [
                `${titleRearrange(insertData.title)} could not be saved to the database.`,
              ],
            });
          } else {
            for (const genreName of data.blockInfo.genres) {
              //check if genre exists
              const [genreRow] = await db
                .select()
                .from(genres)
                .where(eq(genres.genre, genreName));

              const genreId = genreRow.id;

              //insert into genresBooks junction table
              await db.insert(genresBooks).values({
                bookId: book.id,
                genreId: genreId,
              });
            }
            results.push({
              message: `${titleRearrange(insertData.title)} successfully added to database.`,
              actionAttemptItem: {
                ...book,
                genres: data.blockInfo.genres,
                blockID: data.blockID,
              },
              type: data.type,
            });
          }
        }
      }
      return results;
    }),

  edit: adminProcedure
    .input(z.object({ type: mediaType, item: z.unknown() }))
    .mutation(async ({ input, ctx }) => {
      const { type, item } = input;
      const db = ctx.db ?? defaultDb;

      const validation = validate(item, bookCreateSchema);
      if (!validation.valid) {
        return {
          error: 'Schema Violation',
          message: 'Schema violation(s) during edit request',
          errors: validation.errors.map((e) => e.stack),
          actionAttemptItem: item as BookRow,
          type,
        };
      }
      const data = item as BookRow;
      const whereExpr = data.id
        ? eq(books.id, data.id)
        : eq(books.title, data.title);
      const [book] = await db
        .update(books)
        .set({
          title: data.title,
          author: data.author,
          pageCount: data.pageCount ?? null,
          pubYear: data.pubYear ?? null,
          spineColor: data.spineColor,
          imageUrls: data.imageUrls,
        })
        .where(whereExpr)
        .returning();
      if (!book) {
        return {
          error: 'Media Not Found',
          message:
            'Edit requested on an item that does not exist in the database',
          actionAttemptItem: data,
          type,
          errors: [`${data.title} does not exist in the database.`],
        };
      }
      return {
        message: `${titleRearrange(book.title)} successfully edited.`,
        actionAttemptItem: book,
        type,
      };
    }),
});
