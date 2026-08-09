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
  DatabaseSaveEditErrorResponse,
  DatabaseSaveFailureResult,
  DatabaseSaveServerResponse,
  DatabaseSaveSuccessResult,
  PostSavedMediaItem,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { collectedBlockInformationSchema } from '@/mediacollector/collector-form/collectorFormSchema';
import { allGenres, NO_GENRE_FILTER } from '@/lib/enums/genreEnums';
import { DATABASE_SORT_OPTIONS } from 'lib/constants/databaseSortOptions';
import {
  loadBookImagesById,
  loadOtherMediaImagesById,
  replaceBookImageRecords,
  replaceOtherMediaImageRecords,
  resolveAndPersistImageList,
} from 'lib/media-storage/media-image-records';
import type { MediaType } from 'lib/constants/mediaTypes';

const mediaType = z.enum(['book', 'movie', 'videoGame', 'album']);
const sortKey = z.enum(DATABASE_SORT_OPTIONS);
const mediaImageItemSchema = z.object({
  url: z.string().min(1),
  selected: z.boolean().optional(),
  isDefault: z.boolean(),
  spineColor: z.string().min(1),
});

const bookEditSchema = z.object({
  id: z.string().optional(),
  title: z.string().min(1),
  author: z.string().nullable(),
  pageCount: z.number().nullable(),
  pubYear: z.number().nullable(),
  spineColor: z.string().min(1),
  images: z.array(mediaImageItemSchema).min(1),
});
const otherMediaEditSchema = z.object({
  id: z.string().optional(),
  title: z.string().min(1),
  spineColor: z.string().min(1),
  images: z.array(mediaImageItemSchema).min(1),
});

type BookEditItem = z.infer<typeof bookEditSchema>;
type OtherMediaEditItem = z.infer<typeof otherMediaEditSchema>;
type EditableMediaItem = BookEditItem | OtherMediaEditItem;

function resolveDisplaySpineColor(
  images: Array<{ isDefault: boolean; spineColor: string }>,
  fallbackSpineColor: string,
) {
  const defaultImage = images.find((image) => image.isDefault);
  return defaultImage?.spineColor ?? fallbackSpineColor;
}

const IMAGE_SAVE_ROLLED_BACK_CREATION_ERROR =
  'Image failed to save so media item creation was rolled back.';
const IMAGE_SAVE_EDIT_NOT_APPLIED_ERROR =
  'Image failed to save so the edit was not applied.';

function createImagePersistenceErrorResponse(
  title: string,
  rolledBackCreation: boolean,
): DatabaseSaveEditErrorResponse {
  const reason = rolledBackCreation
    ? IMAGE_SAVE_ROLLED_BACK_CREATION_ERROR
    : IMAGE_SAVE_EDIT_NOT_APPLIED_ERROR;
  return {
    title,
    error: 'Image Persistence Error',
    message: reason,
    errors: [reason],
  };
}

function createSaveFailureResult(params: {
  blockID: string;
  title: string;
  error: string;
  message: string;
  errors: string[];
}): DatabaseSaveFailureResult {
  return {
    success: false,
    blockID: params.blockID,
    title: params.title,
    error: params.error,
    message: params.message,
    errors: params.errors,
  };
}

function createSaveImagePersistenceFailure(
  title: string,
  blockID: string,
): DatabaseSaveFailureResult {
  return createSaveFailureResult({
    blockID,
    title,
    error: 'Image Persistence Error',
    message: IMAGE_SAVE_ROLLED_BACK_CREATION_ERROR,
    errors: [IMAGE_SAVE_ROLLED_BACK_CREATION_ERROR],
  });
}

function createSaveSuccessResult(params: {
  blockID: string;
  title: string;
  message: string;
  type: MediaType;
  actionAttemptItem: PostSavedMediaItem & {
    blockID?: string;
    genres?: string[];
  };
}): DatabaseSaveSuccessResult {
  return {
    success: true,
    blockID: params.blockID,
    title: params.title,
    message: params.message,
    type: params.type,
    actionAttemptItem: {
      ...params.actionAttemptItem,
      blockID: params.blockID,
    },
  };
}

function createMediaNotFoundEditResponse(
  data: EditableMediaItem,
  type: z.infer<typeof mediaType>,
) {
  return {
    error: 'Media Not Found' as const,
    message: 'Edit requested on an item that does not exist in the database',
    actionAttemptItem: data,
    type,
    errors: [`${data.title} does not exist in the database.`],
  };
}

export const databaseRouter = router({
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

        const imagesByBookId = await loadBookImagesById(
          db,
          rows.map((row) => row.id),
        );
        const rowsWithResolvedImages = rows.map((row) => {
          const images = imagesByBookId.get(row.id) ?? [];
          return {
            ...row,
            images,
            spineColor: resolveDisplaySpineColor(images, row.spineColor),
          };
        });

        const res: SuccessfulPaginationResponse = {
          message: 'Successful database gather',
          paginatedList: rowsWithResolvedImages,
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

      const imagesByOtherMediaId = await loadOtherMediaImagesById(
        db,
        rows.map((row) => row.id),
      );
      const rowsWithResolvedImages = rows.map((row) => {
        const images = imagesByOtherMediaId.get(row.id) ?? [];
        return {
          ...row,
          images,
          spineColor: resolveDisplaySpineColor(images, row.spineColor),
        };
      });

      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList: rowsWithResolvedImages,
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
              .where(and(eq(otherMedia.mediaType, type), eq(otherMedia.id, id)))
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
        const imagesToPersist =
          selectedImages.length > 0
            ? selectedImages
            : mediaItem.images.length > 0
              ? [{ ...mediaItem.images[0], selected: true, isDefault: true }]
              : [];
        const payload = { ...mediaItem, images: selectedImages };
        payload.images = imagesToPersist.map((image, index) => ({
          ...image,
          isDefault: index === 0,
          selected: index === 0,
        }));
        const validatedData =
          collectedBlockInformationSchema.safeParse(payload);
        if (!validatedData.success) {
          const tree = z.treeifyError(validatedData.error);
          results.push(
            createSaveFailureResult({
              blockID: mediaItem.blockID,
              title: mediaItem.blockInfo.title,
              error: 'Schema Violation',
              message: 'Schema violation(s) during save request',
              errors: tree.errors,
            }),
          );
          continue;
        }

        const data = validatedData.data;
        if (data.type === 'book') {
          let createdBookId: string | null = null;
          try {
            const book = await db.transaction(async (tx) => {
              const [insertedBook] = await tx
                .insert(books)
                .values({
                  title: titleRearrange(data.blockInfo.title),
                  author: data.blockInfo.author ?? '',
                  pageCount: data.blockInfo.pageCount ?? null,
                  pubYear: data.blockInfo.pubYear ?? null,
                  spineColor:
                    data.images[0]?.spineColor ?? data.blockInfo.spineColor,
                })
                .returning();

              if (!insertedBook) {
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
                  bookId: insertedBook.id,
                  genreId: genreRow.id,
                });
              }

              return insertedBook;
            });
            createdBookId = book.id;

            const resolvedImages = await resolveAndPersistImageList(
              { mediaType: 'book', mediaId: book.id },
              data.images.map((image) => ({
                url: image.url,
                isDefault: image.isDefault,
                spineColor: image.spineColor,
              })),
              {
                defaultSpineColor: data.blockInfo.spineColor,
                defaultImageIndex: 0,
              },
            );
            if (resolvedImages.failures.length > 0) {
              await db.delete(books).where(eq(books.id, book.id));
              results.push(
                createSaveImagePersistenceFailure(
                  data.blockInfo.title,
                  data.blockID,
                ),
              );
              continue;
            }
            await replaceBookImageRecords(db, book.id, resolvedImages.images);
            const images = resolvedImages.images.map((image) => ({
              url: image.publicPath,
              isDefault: image.isDefault,
              spineColor: image.spineColor,
              selected: false,
            }));
            const persistedSpineColor = resolveDisplaySpineColor(
              images,
              data.blockInfo.spineColor,
            );

            results.push(
              createSaveSuccessResult({
                blockID: data.blockID,
                title: data.blockInfo.title,
                message: `${titleRearrange(book.title)} successfully added to database.`,
                type: data.type,
                actionAttemptItem: {
                  ...book,
                  images,
                  spineColor: persistedSpineColor,
                  genres: data.blockInfo.genres,
                  blockID: data.blockID,
                },
              }),
            );
          } catch (error) {
            if (createdBookId) {
              await db.delete(books).where(eq(books.id, createdBookId));
            }
            const message =
              error instanceof Error
                ? error.message
                : 'An error occurred while trying to save to the database';
            results.push(
              createSaveFailureResult({
                blockID: data.blockID,
                title: data.blockInfo.title,
                error: 'Database Insertion Error',
                message:
                  'An error occurred while trying to save to the database',
                errors: [message],
              }),
            );
          }
          continue;
        }

        const insertData = {
          mediaType: data.type,
          title: titleRearrange(data.blockInfo.title),
          spineColor: data.images[0]?.spineColor ?? data.blockInfo.spineColor,
        };
        let createdOtherMediaId: string | null = null;
        try {
          const savedOtherMedia = await db.transaction(async (tx) => {
            const [insertedOtherMedia] = await tx
              .insert(otherMedia)
              .values(insertData)
              .returning();
            if (!insertedOtherMedia) {
              throw new Error(
                `${titleRearrange(insertData.title)} could not be saved to the database.`,
              );
            }
            return insertedOtherMedia;
          });
          createdOtherMediaId = savedOtherMedia.id;

          const resolvedImages = await resolveAndPersistImageList(
            { mediaType: data.type, mediaId: savedOtherMedia.id },
            data.images.map((image) => ({
              url: image.url,
              isDefault: image.isDefault,
              spineColor: image.spineColor,
            })),
            {
              defaultSpineColor: data.blockInfo.spineColor,
              defaultImageIndex: 0,
            },
          );
          if (resolvedImages.failures.length > 0) {
            await db
              .delete(otherMedia)
              .where(eq(otherMedia.id, savedOtherMedia.id));
            results.push(
              createSaveImagePersistenceFailure(
                data.blockInfo.title,
                data.blockID,
              ),
            );
            continue;
          }
          await replaceOtherMediaImageRecords(
            db,
            savedOtherMedia.id,
            resolvedImages.images,
          );
          const images = resolvedImages.images.map((image) => ({
            url: image.publicPath,
            isDefault: image.isDefault,
            spineColor: image.spineColor,
            selected: false,
          }));
          const persistedSpineColor = resolveDisplaySpineColor(
            images,
            insertData.spineColor,
          );

          results.push(
            createSaveSuccessResult({
              blockID: data.blockID,
              title: data.blockInfo.title,
              message: `${titleRearrange(insertData.title)} successfully added to database.`,
              type: data.type,
              actionAttemptItem: {
                ...savedOtherMedia,
                images,
                spineColor: persistedSpineColor,
                blockID: data.blockID,
              },
            }),
          );
        } catch (error) {
          if (createdOtherMediaId) {
            await db
              .delete(otherMedia)
              .where(eq(otherMedia.id, createdOtherMediaId));
          }
          const message =
            error instanceof Error
              ? error.message
              : 'An error occurred while trying to save to the database';
          results.push(
            createSaveFailureResult({
              blockID: data.blockID,
              title: data.blockInfo.title,
              error: 'Database Insertion Error',
              message: 'An error occurred while trying to save to the database',
              errors: [message],
            }),
          );
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
        const [existingBook] = await db
          .select()
          .from(books)
          .where(whereExpression)
          .limit(1);
        if (!existingBook) {
          return createMediaNotFoundEditResponse(data, type);
        }
        const resolvedImages = await resolveAndPersistImageList(
          { mediaType: 'book', mediaId: existingBook.id },
          data.images.map((image) => ({
            url: image.url,
            isDefault: image.isDefault,
            spineColor: image.spineColor,
          })),
          { defaultSpineColor: data.spineColor },
        );
        if (resolvedImages.failures.length > 0) {
          return createImagePersistenceErrorResponse(data.title, false);
        }
        const images = resolvedImages.images.map((image) => ({
          url: image.publicPath,
          isDefault: image.isDefault,
          spineColor: image.spineColor,
          selected: false,
        }));
        const persistedSpineColor = resolveDisplaySpineColor(
          images,
          data.spineColor,
        );
        const [book] = await db
          .update(books)
          .set({
            title: data.title,
            author: data.author ?? '',
            pageCount: data.pageCount ?? null,
            pubYear: data.pubYear ?? null,
            spineColor: persistedSpineColor,
          })
          .where(whereExpression)
          .returning();
        if (!book) {
          return createMediaNotFoundEditResponse(data, type);
        }
        await replaceBookImageRecords(db, book.id, resolvedImages.images);

        return {
          message: `${titleRearrange(book.title)} successfully edited.`,
          actionAttemptItem: {
            ...book,
            images,
            spineColor: persistedSpineColor,
          },
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
      const [existingOtherMedia] = await db
        .select()
        .from(otherMedia)
        .where(whereExpression)
        .limit(1);
      if (!existingOtherMedia) {
        return createMediaNotFoundEditResponse(data, type);
      }
      const resolvedImages = await resolveAndPersistImageList(
        { mediaType: type, mediaId: existingOtherMedia.id },
        data.images.map((image) => ({
          url: image.url,
          isDefault: image.isDefault,
          spineColor: image.spineColor,
        })),
        { defaultSpineColor: data.spineColor },
      );
      if (resolvedImages.failures.length > 0) {
        return createImagePersistenceErrorResponse(data.title, false);
      }
      const images = resolvedImages.images.map((image) => ({
        url: image.publicPath,
        isDefault: image.isDefault,
        spineColor: image.spineColor,
        selected: false,
      }));
      const persistedSpineColor = resolveDisplaySpineColor(
        images,
        data.spineColor,
      );
      const [savedOtherMedia] = await db
        .update(otherMedia)
        .set({
          mediaType: type,
          title: data.title,
          spineColor: persistedSpineColor,
        })
        .where(whereExpression)
        .returning();
      if (!savedOtherMedia) {
        return createMediaNotFoundEditResponse(data, type);
      }
      await replaceOtherMediaImageRecords(
        db,
        savedOtherMedia.id,
        resolvedImages.images,
      );

      return {
        message: `${titleRearrange(savedOtherMedia.title)} successfully edited.`,
        actionAttemptItem: {
          ...savedOtherMedia,
          images,
          spineColor: persistedSpineColor,
        },
        type,
      };
    }),
});
