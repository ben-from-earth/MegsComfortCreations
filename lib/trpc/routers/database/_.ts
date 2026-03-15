import { adminProcedure, router } from 'lib/trpc/trpc';
import { TRPCError } from '@trpc/server';
import { and, asc, desc, ilike, isNull, sql, eq } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import {
  books,
  genres,
  genresBooks,
  googleApiQueryUsage,
  otherMedia,
} from '@/db/schema';
import type {
  DatabaseSaveServerResponse,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { searchByTitle } from './actions/search-by-title';
import { collectedBlockInformationSchema } from '@/mediacollector/collector-form/collectorFormSchema';
import { allGenres, NO_GENRE_FILTER } from '@/lib/enums/genreEnums';
import { DATABASE_SORT_OPTIONS } from 'lib/constants/databaseSortOptions';

const mediaType = z.enum(['book', 'movie', 'videoGame', 'album']);
const sortKey = z.enum(DATABASE_SORT_OPTIONS);
const bookEditSchema = z.object({
  id: z.string().optional(),
  title: z.string().min(1),
  author: z.string().nullable(),
  pageCount: z.number().nullable(),
  pubYear: z.number().nullable(),
  spineColor: z.string().min(1),
  imageUrls: z.array(z.string()),
});
const otherMediaEditSchema = z.object({
  id: z.string().optional(),
  title: z.string().min(1),
  spineColor: z.string().min(1),
  imageUrls: z.array(z.string()),
});

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
        sort: sortKey,
        ascDesc: z.enum(['asc', 'desc']),
        genre: z.enum([...allGenres, '', NO_GENRE_FILTER] as const),
      }),
    )
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const { type, limit, page, sort, ascDesc, genre, title } = input;
      const offset = (page - 1) * limit;

      if (type === 'book') {
        const sortColumn = (() => {
          switch (sort) {
            case 'title':
              return books.title;
            case 'author':
              return books.author;
            case 'pageCount':
              return books.pageCount;
            case 'pubYear':
              return books.pubYear;
            default:
              throw new TRPCError({
                code: 'BAD_REQUEST',
                message: `Sort "${sort}" is not supported for books`,
              });
          }
        })();

        const orderExpression =
          ascDesc === 'asc' ? asc(sortColumn) : desc(sortColumn);
        const genreFilter = (() => {
          if (genre === '') {
            return undefined;
          }
          if (genre === NO_GENRE_FILTER) {
            return isNull(genresBooks.bookId);
          }
          return eq(genres.genre, genre);
        })();
        const titleFilter =
          title && title.trim().length > 0
            ? ilike(books.title, `%${title}%`)
            : undefined;
        const whereFilter = and(genreFilter, titleFilter);

        const rows = await db
          .select({
            id: books.id,
            title: books.title,
            author: books.author,
            pageCount: books.pageCount,
            pubYear: books.pubYear,
            spineColor: books.spineColor,
            imageUrls: books.imageUrls,
            mediaType: sql<'book'>`'book'`,
          })
          .from(books)
          .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
          .leftJoin(genres, eq(genresBooks.genreId, genres.id))
          .where(whereFilter)
          .orderBy(orderExpression)
          .limit(limit)
          .offset(offset);

        const [{ count }] = await db
          .select({
            count: sql<number>`cast(count(distinct ${books.id}) as int)`,
          })
          .from(books)
          .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
          .leftJoin(genres, eq(genresBooks.genreId, genres.id))
          .where(whereFilter);

        const res: SuccessfulPaginationResponse = {
          message: 'Successful database gather',
          paginatedList: rows,
          total: count,
        };
        return res;
      }

      if (genre !== '') {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: 'Genre filter is only supported for books',
        });
      }
      if (sort !== 'title') {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: `Sort "${sort}" is not supported for ${type}`,
        });
      }

      const orderExpression =
        ascDesc === 'asc' ? asc(otherMedia.title) : desc(otherMedia.title);
      const titleFilter =
        title && title.trim().length > 0
          ? ilike(otherMedia.title, `%${title}%`)
          : undefined;
      const whereFilter = and(eq(otherMedia.mediaType, type), titleFilter);

      const rows = await db
        .select({
          id: otherMedia.id,
          mediaType: otherMedia.mediaType,
          title: otherMedia.title,
          spineColor: otherMedia.spineColor,
          imageUrls: otherMedia.imageUrls,
        })
        .from(otherMedia)
        .where(whereFilter)
        .orderBy(orderExpression)
        .limit(limit)
        .offset(offset);
      const [{ count }] = await db
        .select({ count: sql<number>`cast(count(*) as int)` })
        .from(otherMedia)
        .where(whereFilter);

      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList: rows,
        total: count,
      };
      return res;
    }),

  delete: adminProcedure
    .input(z.object({ type: mediaType, id: z.string().min(1) }))
    .mutation(async ({ input, ctx }) => {
      const { type, id } = input;
      const db = ctx.db ?? defaultDb;
      const deleted =
        type === 'book'
          ? await db
              .delete(books)
              .where(eq(books.id, id))
              .returning({ id: books.id })
          : await db
              .delete(otherMedia)
              .where(
                and(
                  eq(otherMedia.mediaType, type),
                  eq(otherMedia.id, id),
                ),
              )
              .returning({ id: otherMedia.id });

      if (deleted.length === 0) {
        return {
          message: `No ${type} item with id: ${id} exists`,
        };
      }
      return { message: `Successfully deleted ${type} item` };
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

      for (const mediaItem of input) {
        if (mediaItem.isDatabase) {
          continue;
        }
        const selectedImages = mediaItem.images.filter((img) => img.selected);
        const payload = { ...mediaItem, images: selectedImages };
        const validatedData =
          collectedBlockInformationSchema.safeParse(payload);
        if (!validatedData.success) {
          const tree = z.treeifyError(validatedData.error);
          results.push({
            error: 'Schema Violation',
            message: 'Schema violation(s) during save request',
            errors: tree.errors,
            title: mediaItem.blockInfo.title,
          });
          continue;
        }

        const data = validatedData.data;
        if (data.type === 'book') {
          try {
            const book = await db.transaction(async (tx) => {
              const [createdBook] = await tx
                .insert(books)
                .values({
                  title: titleRearrange(data.blockInfo.title),
                  author: data.blockInfo.author ?? '',
                  pageCount: data.blockInfo.pageCount ?? null,
                  pubYear: data.blockInfo.pubYear ?? null,
                  spineColor: data.blockInfo.spineColor,
                  imageUrls: data.images.map((img) => img.url),
                })
                .returning();

              if (!createdBook) {
                throw new Error('Book insertion failed');
              }

              for (const genreName of data.blockInfo.genres) {
                const [genreRow] = await tx
                  .select()
                  .from(genres)
                  .where(eq(genres.genre, genreName));
                if (!genreRow) {
                  throw new Error(`Genre "${genreName}" does not exist`);
                }
                await tx.insert(genresBooks).values({
                  bookId: createdBook.id,
                  genreId: genreRow.id,
                });
              }
              return createdBook;
            });
            results.push({
              message: `${titleRearrange(book.title)} successfully added to database.`,
              actionAttemptItem: {
                ...book,
                genres: data.blockInfo.genres,
                blockID: data.blockID,
              },
              type: data.type,
            });
          } catch (error) {
            const message =
              error instanceof Error
                ? error.message
                : 'An error occurred while trying to save to the database';
            results.push({
              title: data.blockInfo.title,
              error: 'Database Insertion Error',
              message: 'An error occurred while trying to save to the database',
              errors: [message],
            });
          }
          continue;
        }

        const insertData = {
          mediaType: data.type,
          title: titleRearrange(data.blockInfo.title),
          spineColor: data.blockInfo.spineColor,
          imageUrls: data.images.map((img) => img.url),
        };
        const [savedOtherMedia] = await db
          .insert(otherMedia)
          .values(insertData)
          .returning();

        if (!savedOtherMedia) {
          results.push({
            title: data.blockInfo.title,
            error: 'Database Insertion Error',
            message: 'An error occurred while trying to save to the database',
            errors: [
              `${titleRearrange(insertData.title)} could not be saved to the database.`,
            ],
          });
        } else {
          results.push({
            message: `${titleRearrange(insertData.title)} successfully added to database.`,
            actionAttemptItem: {
              ...savedOtherMedia,
              blockID: data.blockID,
            },
            type: data.type,
          });
        }
      }
      return results;
    }),

  edit: adminProcedure
    .input(z.object({ type: mediaType, item: z.unknown() }))
    .mutation(async ({ input, ctx }) => {
      const { type, item } = input;
      const db = ctx.db ?? defaultDb;

      if (type === 'book') {
        const validation = bookEditSchema.safeParse(item);
        if (!validation.success) {
          return {
            error: 'Schema Violation',
            message: 'Schema violation(s) during edit request',
            errors: validation.error.issues.map((issue) => issue.message),
            actionAttemptItem: item,
            type,
          };
        }
        const data = validation.data;
        const whereExpression = data.id
          ? eq(books.id, data.id)
          : eq(books.title, data.title);
        const [book] = await db
          .update(books)
          .set({
            title: data.title,
            author: data.author ?? '',
            pageCount: data.pageCount ?? null,
            pubYear: data.pubYear ?? null,
            spineColor: data.spineColor,
            imageUrls: data.imageUrls,
          })
          .where(whereExpression)
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
      }

      const validation = otherMediaEditSchema.safeParse(item);
      if (!validation.success) {
        return {
          error: 'Schema Violation',
          message: 'Schema violation(s) during edit request',
          errors: validation.error.issues.map((issue) => issue.message),
          actionAttemptItem: item,
          type,
        };
      }
      const data = validation.data;
      const whereExpression = data.id
        ? and(eq(otherMedia.id, data.id), eq(otherMedia.mediaType, type))
        : and(eq(otherMedia.title, data.title), eq(otherMedia.mediaType, type));
      const [savedOtherMedia] = await db
        .update(otherMedia)
        .set({
          mediaType: type,
          title: data.title,
          spineColor: data.spineColor,
          imageUrls: data.imageUrls,
        })
        .where(whereExpression)
        .returning();
      if (!savedOtherMedia) {
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
        message: `${titleRearrange(savedOtherMedia.title)} successfully edited.`,
        actionAttemptItem: savedOtherMedia,
        type,
      };
    }),
});
