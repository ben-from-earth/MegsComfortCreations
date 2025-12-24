import { publicProcedure, router } from '@/lib/trpc/trpc';
import { asc, desc, ilike, sql, eq } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/app/db/client';
import { albums, books, movies, videoGames } from '@/app/db/schema';
import type {
  SuccessfulPaginationResponse,
  SuccessfulMediaSaveEditResponse,
  SuccessfulMediaSearchResponse,
  BookInsert,
  MovieInsert,
  VideoGameInsert,
  AlbumInsert,
  PostSavedMediaItem,
  BookRow,
  MovieRow,
  VideoGameRow,
  AlbumRow,
} from '@/lib/interfaces/globalInterfaces';
import { validate } from 'jsonschema';
import bookCreateSchema from '@/lib/database/schemas/bookCreateSchema.json';
import otherMediaCreateSchema from '@/lib/database/schemas/otherMediaCreateSchema.json';
import { titleRearrange } from '@/lib/helpers/titleRearrange';

const mediaType = z.enum(['book', 'movie', 'video_game', 'album']);

const tableMap = {
  book: books,
  movie: movies,
  video_game: videoGames,
  album: albums,
} as const;

export const databaseRouter = router({
  searchByTitle: publicProcedure
    .input(z.object({ type: mediaType, title: z.string().min(1) }))
    .query(async ({ input, ctx }) => {
      console.log('Database searchByTitle called with input:', input);
      const db = ctx.db ?? defaultDb;
      const { type, title } = input;
      const table = tableMap[type];

      const rearrangedTitle = titleRearrange(title);
      const result = await db
        .select()
        .from(table)
        .where(ilike(table.title, rearrangedTitle));
      const total = result.length;

      if (total === 0) {
        return {
          error: 'Media Not Found',
          message: `No ${type} in database with title ${rearrangedTitle}`,
          failedSearchData: [],
        };
      }

      return {
        message: `Successfully found ${total} ${type}(s) with title ${titleRearrange(
          result[0].title,
        )}`,
        foundMediaList: result,
        total,
      } satisfies SuccessfulMediaSearchResponse;
    }),
  getPaginated: publicProcedure
    .input(
      z.object({
        type: mediaType,
        limit: z.number().int().positive(),
        page: z.number().int().positive(),
        sort: z.enum(['title', 'author', 'pubYear', 'spineColor']),
        ascDesc: z.enum(['asc', 'desc']),
      }),
    )
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const { type, limit, page, sort, ascDesc } = input;
      const table = tableMap[type];
      const offset = (page - 1) * limit;

      // Determine the correct sort column per media type
      const sortColumn = (() => {
        if (type === 'book') {
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
        }
        if (type === 'movie') {
          return sort === 'spineColor' ? movies.spineColor : movies.title;
        }
        if (type === 'video_game') {
          return sort === 'spineColor'
            ? videoGames.spineColor
            : videoGames.title;
        }
        return sort === 'spineColor' ? albums.spineColor : albums.title;
      })();

      const orderExpr = ascDesc === 'asc' ? asc(sortColumn) : desc(sortColumn);

      const rows = await db
        .select()
        .from(table)
        .orderBy(orderExpr)
        .limit(limit)
        .offset(offset);

      const [{ count }] = await db
        .select({ count: sql<number>`cast(count(*) as int)` })
        .from(table);

      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList: rows,
        total: count,
      };
      return res;
    }),

  deleteByTitle: publicProcedure
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

  save: publicProcedure
    .input(z.object({ type: mediaType, mediaData: z.unknown() }))
    .mutation(async ({ input, ctx }) => {
      const { type, mediaData } = input;
      const db = ctx.db ?? defaultDb;

      if (type === 'book') {
        const validation = validate(mediaData, bookCreateSchema);
        if (!validation.valid) {
          return {
            error: 'Schema Violation',
            message: 'Schema violation(s) during save request',
            errors: validation.errors.map((e) => e.stack),
            actionAttemptItem: mediaData as BookInsert,
            type,
          };
        }
        const [book] = await db
          .insert(books)
          .values({
            title: (mediaData as BookInsert).title,
            author: (mediaData as BookInsert).author,
            pageCount: (mediaData as BookInsert).pageCount ?? null,
            pubYear: (mediaData as BookInsert).pubYear ?? null,
            spineColor: (mediaData as BookInsert).spineColor,
            imageUrls: (mediaData as BookInsert).imageUrls,
          })
          .returning();
        const res: SuccessfulMediaSaveEditResponse = {
          message: `${titleRearrange((mediaData as BookInsert).title)} successfully added to database.`,
          actionAttemptItem: {
            ...book,
            genres: (mediaData as BookInsert).genres,
            blockID: (mediaData as BookInsert).blockID,
          },
          type,
        };
        return res;
      }

      const validation = validate(mediaData, otherMediaCreateSchema);
      if (!validation.valid) {
        return {
          error: 'Schema Violation',
          message: 'Schema violation(s) during save request',
          errors: validation.errors.map((e) => e.stack),
          actionAttemptItem: mediaData as PostSavedMediaItem,
          type,
        };
      }

      // Non-book saves: validate and insert with precise types per media
      switch (type) {
        case 'movie': {
          const data = mediaData as MovieInsert;
          const [row] = await db
            .insert(movies)
            .values({
              title: data.title,
              spineColor: data.spineColor,
              imageUrls: data.imageUrls,
            })
            .returning();
          return {
            message: `${titleRearrange(data.title)} successfully added to database.`,
            actionAttemptItem: { ...row, blockID: data.blockID },
            type,
          } satisfies SuccessfulMediaSaveEditResponse;
        }
        case 'video_game': {
          const data = mediaData as VideoGameInsert;
          const [row] = await db
            .insert(videoGames)
            .values({
              title: data.title,
              spineColor: data.spineColor,
              imageUrls: data.imageUrls,
            })
            .returning();
          return {
            message: `${titleRearrange(data.title)} successfully added to database.`,
            actionAttemptItem: { ...row, blockID: data.blockID },
            type,
          } satisfies SuccessfulMediaSaveEditResponse;
        }
        case 'album': {
          const data = mediaData as AlbumInsert;
          const [row] = await db
            .insert(albums)
            .values({
              title: data.title,
              spineColor: data.spineColor,
              imageUrls: data.imageUrls,
            })
            .returning();
          return {
            message: `${titleRearrange(data.title)} successfully added to database.`,
            actionAttemptItem: { ...row, blockID: data.blockID },
            type,
          } satisfies SuccessfulMediaSaveEditResponse;
        }
      }
    }),

  edit: publicProcedure
    .input(z.object({ type: mediaType, item: z.unknown() }))
    .mutation(async ({ input, ctx }) => {
      const { type, item } = input;
      const db = ctx.db ?? defaultDb;

      switch (type) {
        case 'book': {
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
          } satisfies SuccessfulMediaSaveEditResponse;
        }
        case 'movie': {
          const validation = validate(item, otherMediaCreateSchema);
          if (!validation.valid) {
            return {
              error: 'Schema Violation',
              message: 'Schema violation(s) during edit request',
              errors: validation.errors.map((e) => e.stack),
              actionAttemptItem: item as PostSavedMediaItem,
              type,
            };
          }
          const data = item as MovieRow;
          const whereExpr = data.id
            ? eq(movies.id, data.id)
            : eq(movies.title, data.title);
          const [row] = await db
            .update(movies)
            .set({
              title: data.title,
              spineColor: data.spineColor,
              imageUrls: data.imageUrls,
            })
            .where(whereExpr)
            .returning();
          if (!row) {
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
            message: `${titleRearrange(row.title)} successfully edited.`,
            actionAttemptItem: row,
            type,
          } satisfies SuccessfulMediaSaveEditResponse;
        }
        case 'video_game': {
          const validation = validate(item, otherMediaCreateSchema);
          if (!validation.valid) {
            return {
              error: 'Schema Violation',
              message: 'Schema violation(s) during edit request',
              errors: validation.errors.map((e) => e.stack),
              actionAttemptItem: item as PostSavedMediaItem,
              type,
            };
          }
          const data = item as VideoGameRow;
          const whereExpr = data.id
            ? eq(videoGames.id, data.id)
            : eq(videoGames.title, data.title);
          const [row] = await db
            .update(videoGames)
            .set({
              title: data.title,
              spineColor: data.spineColor,
              imageUrls: data.imageUrls,
            })
            .where(whereExpr)
            .returning();
          if (!row) {
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
            message: `${titleRearrange(row.title)} successfully edited.`,
            actionAttemptItem: row,
            type,
          } satisfies SuccessfulMediaSaveEditResponse;
        }
        case 'album': {
          const validation = validate(item, otherMediaCreateSchema);
          if (!validation.valid) {
            return {
              error: 'Schema Violation',
              message: 'Schema violation(s) during edit request',
              errors: validation.errors.map((e) => e.stack),
              actionAttemptItem: item as PostSavedMediaItem,
              type,
            };
          }
          const data = item as AlbumRow;
          const whereExpr = data.id
            ? eq(albums.id, data.id)
            : eq(albums.title, data.title);
          const [row] = await db
            .update(albums)
            .set({
              title: data.title,
              spineColor: data.spineColor,
              imageUrls: data.imageUrls,
            })
            .where(whereExpr)
            .returning();
          if (!row) {
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
            message: `${titleRearrange(row.title)} successfully edited.`,
            actionAttemptItem: row,
            type,
          } satisfies SuccessfulMediaSaveEditResponse;
        }
      }
    }),
});
