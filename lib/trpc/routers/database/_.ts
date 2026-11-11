import { adminProcedure, router } from 'lib/trpc/trpc';
import { TRPCError } from '@trpc/server';
import {
  and,
  asc,
  desc,
  eq,
  exists,
  ilike,
  inArray,
  notExists,
  sql,
} from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb, type Db } from '@/db/client';
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
  DatabaseSaveEditErrorResponse,
  DatabaseSaveFailureResult,
  DatabaseSaveServerResponse,
  DatabaseSaveSuccessResult,
  MediaImageItem,
  PostSavedMediaItem,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/global-interfaces';
import {
  splitMovieTitle,
  toDisplayMovieTitle,
  toStoredMovieTitle,
} from 'lib/helpers/movie-title';
import {
  toDisplayTitle,
  toStoredTitle,
  toTitleSearchTerm,
} from 'lib/helpers/title-articles';
import {
  mediaItemFormSchema,
  PLACEHOLDER_MEDIA_IMAGE_URL,
} from '@/mediacollector/collector-form/media-item-form-schema';
import { allGenres, NO_GENRE_FILTER } from '@/lib/enums/genre-enums';
import { DATABASE_SORT_OPTIONS } from 'lib/constants/database-sort-options';
import {
  MEDIA_TYPE_DEFINITIONS,
  type MediaType,
} from 'lib/constants/media-types';
import { MOVIE_FORMATS, MOVIE_FORMAT_OPTIONS } from 'lib/constants/movie-formats';
import {
  appendBookCoverRecords,
  coversToPersistForNewItem,
  listCoversToAppend,
  loadAlbumImagesById,
  loadBookImagesById,
  loadMovieImagesById,
  loadVideoGameImagesById,
  replaceBookImageRecords,
  replaceMediaImageRecords,
  resolveAndPersistImageList,
} from 'lib/media-storage/media-image-records';

const mediaType = z.enum(['book', 'movie', 'videoGame', 'album']);
const sortKey = z.enum(DATABASE_SORT_OPTIONS);
const mediaImageItemSchema = z.object({
  url: z.string().min(1),
  selected: z.boolean().optional(),
  isDefault: z.boolean(),
  spineColor: z.string().min(1),
  codeColor: z.string().nullable().optional(),
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
const titleMediaEditSchema = z.object({
  id: z.string().optional(),
  title: z.string().min(1),
  spineColor: z.string().min(1),
  images: z.array(mediaImageItemSchema).min(1),
  movieFormat: z.enum(MOVIE_FORMATS).optional(),
});

const albumEditSchema = z.object({
  id: z.string().optional(),
  title: z.string().min(1),
  artist: z.string().nullable(),
  releaseDate: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/)
    .nullable()
    .optional(),
  spotifyAlbumUri: z.string().min(1).optional(),
  spineColor: z.string().min(1),
  images: z.array(mediaImageItemSchema).min(1),
});

type BookEditItem = z.infer<typeof bookEditSchema>;
type TitleMediaEditItem = z.infer<typeof titleMediaEditSchema>;
type AlbumEditItem = z.infer<typeof albumEditSchema>;
type EditableMediaItem = BookEditItem | TitleMediaEditItem | AlbumEditItem;

const ALBUM_ROW_SPINE_COLOR = '#ffffff';

function movieLibraryTitleFilter(title: string) {
  const { baseTitle, year } = splitMovieTitle(title);
  const containsTitle = ilike(
    movies.title,
    `%${toTitleSearchTerm(baseTitle)}%`,
  );
  if (!year) {
    return containsTitle;
  }
  return and(containsTitle, ilike(movies.title, `% (${year})`));
}

function databaseErrorHint(error: unknown) {
  const parts: string[] = [];
  const visit = (value: unknown, depth: number) => {
    if (depth > 4 || value === null || value === undefined) {
      return;
    }
    if (value instanceof Error) {
      parts.push(value.message);
      visit(value.cause, depth + 1);
    }
    if (typeof value === 'object' && value !== null && 'constraint' in value) {
      const constraint = value.constraint;
      if (typeof constraint === 'string') {
        parts.push(constraint);
      }
    }
  };
  visit(error, 0);
  return parts.join(' ');
}

function mediaTypeLabel(type: MediaType) {
  return MEDIA_TYPE_DEFINITIONS.find(
    (definition) => definition.mediaType === type,
  )?.label;
}

function savedImageList(
  images: Awaited<ReturnType<typeof resolveAndPersistImageList>>['images'],
) {
  return images.map((image) => ({
    url: image.publicPath,
    isDefault: image.isDefault,
    spineColor: image.spineColor,
    codeColor: image.codeColor,
    selected: false,
  }));
}

function toPostSavedMediaItem(
  row:
    | typeof movies.$inferSelect
    | typeof videoGames.$inferSelect
    | typeof albums.$inferSelect,
  presentation: {
    title: string;
    images: MediaImageItem[];
    spineColor: string;
  },
): PostSavedMediaItem {
  if ('movieFormat' in row) {
    return {
      mediaType: 'movie',
      id: row.id,
      title: presentation.title,
      spineColor: presentation.spineColor,
      images: presentation.images,
      movieFormat: row.movieFormat,
    };
  }

  if ('artist' in row) {
    return {
      mediaType: 'album',
      id: row.id,
      title: presentation.title,
      spineColor: presentation.spineColor,
      images: presentation.images,
      artist: row.artist,
      releaseDate: row.releaseDate,
      spotifyAlbumUri: row.spotifyAlbumUri,
    };
  }

  return {
    mediaType: 'videoGame',
    id: row.id,
    title: presentation.title,
    spineColor: presentation.spineColor,
    images: presentation.images,
  };
}

async function persistCreatedImages(
  db: Db,
  mediaType: Exclude<MediaType, 'book'>,
  mediaId: string,
  sourceImages: Array<{
    url: string;
    isDefault: boolean;
    spineColor: string;
    codeColor?: string | null;
  }>,
  defaultSpineColor: string,
) {
  const resolvedImages = await resolveAndPersistImageList(
    { mediaType, mediaId },
    sourceImages,
    { defaultSpineColor, defaultImageIndex: 0 },
  );
  if (resolvedImages.failures.length > 0) {
    return null;
  }
  await replaceMediaImageRecords(
    db,
    { mediaType, mediaId },
    resolvedImages.images,
  );
  return savedImageList(resolvedImages.images);
}

async function deleteSavedMedia(
  db: Db,
  type: Exclude<MediaType, 'book'>,
  id: string,
) {
  if (type === 'movie') {
    await db.delete(movies).where(eq(movies.id, id));
    return;
  }
  if (type === 'videoGame') {
    await db.delete(videoGames).where(eq(videoGames.id, id));
    return;
  }
  await db.delete(albums).where(eq(albums.id, id));
}

function resolveDisplaySpineColor(
  images: Array<{ isDefault: boolean; spineColor: string }>,
  fallbackSpineColor: string,
) {
  const defaultImage = images.find((image) => image.isDefault);
  return defaultImage?.spineColor ?? fallbackSpineColor;
}

async function loadBookGenresById(db: Db, bookIds: string[]) {
  if (bookIds.length === 0) {
    return new Map<string, string[]>();
  }

  const rows = await db
    .select({
      bookId: genresBooks.bookId,
      genre: genres.genre,
    })
    .from(genresBooks)
    .innerJoin(genres, eq(genresBooks.genreId, genres.id))
    .where(inArray(genresBooks.bookId, bookIds));

  const genresByBookId = new Map<string, string[]>();
  for (const row of rows) {
    const currentGenres = genresByBookId.get(row.bookId) ?? [];
    currentGenres.push(row.genre);
    genresByBookId.set(row.bookId, currentGenres);
  }
  return genresByBookId;
}

const IMAGE_SAVE_ROLLED_BACK_CREATION_ERROR =
  'Image failed to save so media item creation was rolled back.';
const IMAGE_SAVE_EDIT_NOT_APPLIED_ERROR =
  'Image failed to save so the edit was not applied.';
const COVER_APPEND_FAILED_MESSAGE =
  'The new cover could not be added to this database item.';

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
        limit: z.union([z.literal(12), z.literal(24), z.literal(48)]),
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
            return notExists(
              db
                .select({ bookId: genresBooks.bookId })
                .from(genresBooks)
                .where(eq(genresBooks.bookId, books.id)),
            );
          }
          return exists(
            db
              .select({ bookId: genresBooks.bookId })
              .from(genresBooks)
              .innerJoin(genres, eq(genresBooks.genreId, genres.id))
              .where(
                and(eq(genresBooks.bookId, books.id), eq(genres.genre, genre)),
              ),
          );
        })();
        const titleFilter =
          title && title.trim().length > 0
            ? ilike(books.title, `%${toTitleSearchTerm(title)}%`)
            : undefined;
        const whereFilter = and(genreFilter, titleFilter);

        const [rows, [{ count }]] = await Promise.all([
          db
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
            .where(whereFilter)
            .orderBy(orderExpression)
            .limit(limit)
            .offset(offset),
          db
            .select({
              count: sql<number>`cast(count(*) as int)`,
            })
            .from(books)
            .where(whereFilter),
        ]);
        const bookIds = rows.map((row) => row.id);
        const [imagesByBookId, genresByBookId] = await Promise.all([
          loadBookImagesById(db, bookIds),
          loadBookGenresById(db, bookIds),
        ]);

        const rowsWithResolvedImages = rows.map((row) => {
          const images = imagesByBookId.get(row.id) ?? [];
          return {
            ...row,
            title: toDisplayTitle(row.title),
            images,
            spineColor: resolveDisplaySpineColor(images, row.spineColor),
            genres: genresByBookId.get(row.id) ?? [],
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

      const hasTitle = title !== undefined && title.trim().length > 0;
      const titleContains = hasTitle
        ? `%${toTitleSearchTerm(title)}%`
        : undefined;

      if (type === 'movie') {
        const titleFilter = hasTitle ? movieLibraryTitleFilter(title) : undefined;
        const orderExpression =
          ascDesc === 'asc' ? asc(movies.title) : desc(movies.title);
        const [rows, [{ count }]] = await Promise.all([
          db
            .select({
              id: movies.id,
              title: movies.title,
              spineColor: movies.spineColor,
              movieFormat: movies.movieFormat,
              mediaType: sql<'movie'>`'movie'`,
            })
            .from(movies)
            .where(titleFilter)
            .orderBy(orderExpression)
            .limit(limit)
            .offset(offset),
          db
            .select({ count: sql<number>`cast(count(*) as int)` })
            .from(movies)
            .where(titleFilter),
        ]);
        const imagesById = await loadMovieImagesById(
          db,
          rows.map((row) => row.id),
        );
        const rowsWithResolvedImages = rows.map((row) => {
          const images = imagesById.get(row.id) ?? [];
          return {
            ...row,
            title: toDisplayMovieTitle(row.title),
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

      if (type === 'videoGame') {
        const titleFilter = titleContains
          ? ilike(videoGames.title, titleContains)
          : undefined;
        const orderExpression =
          ascDesc === 'asc' ? asc(videoGames.title) : desc(videoGames.title);
        const [rows, [{ count }]] = await Promise.all([
          db
            .select({
              id: videoGames.id,
              title: videoGames.title,
              spineColor: videoGames.spineColor,
              mediaType: sql<'videoGame'>`'videoGame'`,
            })
            .from(videoGames)
            .where(titleFilter)
            .orderBy(orderExpression)
            .limit(limit)
            .offset(offset),
          db
            .select({ count: sql<number>`cast(count(*) as int)` })
            .from(videoGames)
            .where(titleFilter),
        ]);
        const imagesById = await loadVideoGameImagesById(
          db,
          rows.map((row) => row.id),
        );
        const rowsWithResolvedImages = rows.map((row) => {
          const images = imagesById.get(row.id) ?? [];
          return {
            ...row,
            title: toDisplayTitle(row.title),
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

      const titleFilter = titleContains
        ? ilike(albums.title, titleContains)
        : undefined;
      const orderExpression =
        ascDesc === 'asc' ? asc(albums.title) : desc(albums.title);
      const [rows, [{ count }]] = await Promise.all([
        db
          .select({
            id: albums.id,
            title: albums.title,
            artist: albums.artist,
            releaseDate: albums.releaseDate,
            spotifyAlbumUri: albums.spotifyAlbumUri,
            spineColor: albums.spineColor,
            mediaType: sql<'album'>`'album'`,
          })
          .from(albums)
          .where(titleFilter)
          .orderBy(orderExpression)
          .limit(limit)
          .offset(offset),
        db
          .select({ count: sql<number>`cast(count(*) as int)` })
          .from(albums)
          .where(titleFilter),
      ]);
      const imagesById = await loadAlbumImagesById(
        db,
        rows.map((row) => row.id),
      );
      const rowsWithResolvedImages = rows.map((row) => {
        const images = imagesById.get(row.id) ?? [];
        return {
          ...row,
          title: toDisplayTitle(row.title),
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
          : type === 'movie'
            ? await db
                .delete(movies)
                .where(eq(movies.id, id))
                .returning({ id: movies.id })
            : type === 'videoGame'
              ? await db
                  .delete(videoGames)
                  .where(eq(videoGames.id, id))
                  .returning({ id: videoGames.id })
              : await db
                  .delete(albums)
                  .where(eq(albums.id, id))
                  .returning({ id: albums.id });

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
    .input(z.array(mediaItemFormSchema))
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;

      const results: DatabaseSaveServerResponse = [];

      for (const mediaItem of input) {
        if (mediaItem.isDatabase) {
          const uploadedCovers =
            mediaItem.type === 'book'
              ? mediaItem.images.filter(
                  (image) =>
                    image.uploaded === true &&
                    image.url !== PLACEHOLDER_MEDIA_IMAGE_URL,
                )
              : [];
          if (uploadedCovers.length > 0) {
            try {
              const imagesByBookId = await loadBookImagesById(db, [
                mediaItem.blockID,
              ]);
              const storedImages = imagesByBookId.get(mediaItem.blockID) ?? [];
              const coversToAppend = listCoversToAppend(
                storedImages,
                uploadedCovers,
              );
              if (coversToAppend.length > 0) {
                await appendBookCoverRecords(
                  db,
                  mediaItem.blockID,
                  coversToAppend,
                );
              }
            } catch (error) {
              const message =
                error instanceof Error
                  ? error.message
                  : COVER_APPEND_FAILED_MESSAGE;
              results.push(
                createSaveFailureResult({
                  blockID: mediaItem.blockID,
                  title: mediaItem.blockInfo.title,
                  error: 'Cover Append Error',
                  message: COVER_APPEND_FAILED_MESSAGE,
                  errors: [message],
                }),
              );
            }
          }
          continue;
        }
        const payload = {
          ...mediaItem,
          images: coversToPersistForNewItem(mediaItem.images),
        };
        const validatedData = mediaItemFormSchema.safeParse(payload);
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
                  title: toStoredTitle(data.blockInfo.title),
                  author: data.blockInfo.author ?? '',
                  pageCount: data.blockInfo.pageCount ?? null,
                  pubYear: data.blockInfo.pubYear ?? null,
                  spineColor: resolveDisplaySpineColor(
                    data.images,
                    data.blockInfo.spineColor,
                  ),
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
                message: `${toDisplayTitle(book.title)} successfully added to database.`,
                type: data.type,
                actionAttemptItem: {
                  mediaType: 'book',
                  id: book.id,
                  title: toDisplayTitle(book.title),
                  author: book.author,
                  pageCount: book.pageCount,
                  pubYear: book.pubYear,
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
            const duplicateHint = `${error instanceof Error ? error.message : ''} ${
              error instanceof Error && error.cause instanceof Error
                ? error.cause.message
                : ''
            }`;
            if (duplicateHint.includes('books_title_author_unique')) {
              results.push(
                createSaveFailureResult({
                  blockID: data.blockID,
                  title: data.blockInfo.title,
                  error: 'Duplicate Book',
                  message: 'A book with this title and author already exists.',
                  errors: ['Duplicate Book'],
                }),
              );
              continue;
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

        if (data.type === 'album' && data.images.length === 0) {
          results.push(
            createSaveFailureResult({
              blockID: data.blockID,
              title: data.blockInfo.title,
              error: 'Schema Violation',
              message: 'Cover image is Required',
              errors: ['Cover image is Required'],
            }),
          );
          continue;
        }

        const displayTitle =
          data.type === 'movie'
            ? toDisplayMovieTitle(toStoredMovieTitle(data.blockInfo.title))
            : toDisplayTitle(toStoredTitle(data.blockInfo.title));
        let createdMediaId: string | null = null;
        try {
          const savedMedia =
            data.type === 'movie'
              ? await db.transaction(async (tx) => {
                  const [insertedMovie] = await tx
                    .insert(movies)
                    .values({
                      title: toStoredMovieTitle(data.blockInfo.title),
                      spineColor: resolveDisplaySpineColor(
                        data.images,
                        data.blockInfo.spineColor,
                      ),
                      movieFormat: data.blockInfo.movieFormat ?? 'dvd',
                    })
                    .returning();
                  if (!insertedMovie) {
                    throw new Error(
                      `${displayTitle} could not be saved to the database.`,
                    );
                  }
                  return insertedMovie;
                })
              : data.type === 'videoGame'
                ? await db.transaction(async (tx) => {
                    const [insertedVideoGame] = await tx
                      .insert(videoGames)
                      .values({
                        title: toStoredTitle(data.blockInfo.title),
                        spineColor: resolveDisplaySpineColor(
                          data.images,
                          data.blockInfo.spineColor,
                        ),
                      })
                      .returning();
                    if (!insertedVideoGame) {
                      throw new Error(
                        `${displayTitle} could not be saved to the database.`,
                      );
                    }
                    return insertedVideoGame;
                  })
                : await db.transaction(async (tx) => {
                    const [insertedAlbum] = await tx
                      .insert(albums)
                      .values({
                        title: toStoredTitle(data.blockInfo.title),
                        artist: data.blockInfo.artist?.trim() ?? '',
                        releaseDate: data.blockInfo.releaseDate ?? null,
                        spotifyAlbumUri:
                          data.blockInfo.spotifyAlbumUri?.trim() || null,
                        spineColor: ALBUM_ROW_SPINE_COLOR,
                      })
                      .returning();
                    if (!insertedAlbum) {
                      throw new Error(
                        `${displayTitle} could not be saved to the database.`,
                      );
                    }
                    return insertedAlbum;
                  });
          createdMediaId = savedMedia.id;

          const sourceImages =
            data.type === 'album'
              ? data.images.map((image) => ({
                  url: image.url,
                  isDefault: image.isDefault,
                  spineColor: ALBUM_ROW_SPINE_COLOR,
                  codeColor: image.codeColor ?? null,
                }))
              : data.images.map((image) => ({
                  url: image.url,
                  isDefault: image.isDefault,
                  spineColor: image.spineColor,
                }));
          const images = await persistCreatedImages(
            db,
            data.type,
            savedMedia.id,
            sourceImages,
            data.type === 'album'
              ? ALBUM_ROW_SPINE_COLOR
              : data.blockInfo.spineColor,
          );
          if (!images) {
            await deleteSavedMedia(db, data.type, savedMedia.id);
            results.push(
              createSaveImagePersistenceFailure(
                data.blockInfo.title,
                data.blockID,
              ),
            );
            continue;
          }
          const persistedSpineColor =
            data.type === 'album'
              ? ALBUM_ROW_SPINE_COLOR
              : resolveDisplaySpineColor(images, data.blockInfo.spineColor);

          results.push(
            createSaveSuccessResult({
              blockID: data.blockID,
              title: data.blockInfo.title,
              message: `${displayTitle} successfully added to database.`,
              type: data.type,
              actionAttemptItem: toPostSavedMediaItem(savedMedia, {
                title: displayTitle,
                images,
                spineColor: persistedSpineColor,
              }),
            }),
          );
        } catch (error) {
          if (createdMediaId) {
            await deleteSavedMedia(db, data.type, createdMediaId);
          }
          const duplicateHint = databaseErrorHint(error);
          const label = mediaTypeLabel(data.type);
          const isDuplicateMovie =
            data.type === 'movie' &&
            duplicateHint.includes('movies_title_format_unique');
          const isDuplicateVideoGame =
            data.type === 'videoGame' &&
            duplicateHint.includes('video_games_title_unique');
          const isDuplicateAlbum =
            data.type === 'album' &&
            duplicateHint.includes('albums_title_artist_unique');
          if (label && (isDuplicateMovie || isDuplicateVideoGame || isDuplicateAlbum)) {
            const formatLabel = MOVIE_FORMAT_OPTIONS.find(
              (option) =>
                option.value === (data.blockInfo.movieFormat ?? 'dvd'),
            )?.label;
            const formatSuffix =
              isDuplicateMovie && formatLabel ? ` (${formatLabel})` : '';
            const artistSuffix =
              isDuplicateAlbum && data.blockInfo.artist?.trim()
                ? ` by ${data.blockInfo.artist.trim()}`
                : '';
            results.push(
              createSaveFailureResult({
                blockID: data.blockID,
                title: data.blockInfo.title,
                error: `Duplicate ${label}`,
                message: `${label} "${displayTitle}"${artistSuffix}${formatSuffix} is already stored.`,
                errors: [`Duplicate ${label}`],
              }),
            );
            continue;
          }
          if (
            data.type === 'album' &&
            duplicateHint.includes('albums_spotify_album_uri_unique')
          ) {
            const uri = data.blockInfo.spotifyAlbumUri?.trim();
            const [existingAlbum] = uri
              ? await db
                  .select()
                  .from(albums)
                  .where(eq(albums.spotifyAlbumUri, uri))
                  .limit(1)
              : [];
            if (existingAlbum) {
              const imagesById = await loadAlbumImagesById(db, [
                existingAlbum.id,
              ]);
              const images = imagesById.get(existingAlbum.id) ?? [];
              results.push(
                createSaveSuccessResult({
                  blockID: data.blockID,
                  title: data.blockInfo.title,
                  message: `${toDisplayTitle(existingAlbum.title)} is already stored.`,
                  type: 'album',
                  actionAttemptItem: toPostSavedMediaItem(existingAlbum, {
                    title: toDisplayTitle(existingAlbum.title),
                    images,
                    spineColor: resolveDisplaySpineColor(
                      images,
                      existingAlbum.spineColor,
                    ),
                  }),
                }),
              );
              continue;
            }
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
          : eq(books.title, toStoredTitle(data.title));
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
            title: toStoredTitle(data.title),
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
          message: `${toDisplayTitle(book.title)} successfully edited.`,
          actionAttemptItem: {
            mediaType: 'book',
            id: book.id,
            title: toDisplayTitle(book.title),
            author: book.author,
            pageCount: book.pageCount,
            pubYear: book.pubYear,
            images,
            spineColor: persistedSpineColor,
          },
          type,
        };
      }

      if (type === 'album') {
        const validation = albumEditSchema.safeParse(item);
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
          ? eq(albums.id, data.id)
          : and(
              eq(albums.title, toStoredTitle(data.title)),
              eq(albums.artist, data.artist ?? ''),
            );
        const [existingAlbum] = await db
          .select()
          .from(albums)
          .where(whereExpression)
          .limit(1);
        if (!existingAlbum) {
          return createMediaNotFoundEditResponse(data, type);
        }
        const resolvedImages = await resolveAndPersistImageList(
          { mediaType: 'album', mediaId: existingAlbum.id },
          data.images.map((image) => ({
            url: image.url,
            isDefault: image.isDefault,
            spineColor: ALBUM_ROW_SPINE_COLOR,
            codeColor: image.codeColor ?? null,
          })),
          { defaultSpineColor: ALBUM_ROW_SPINE_COLOR },
        );
        if (resolvedImages.failures.length > 0) {
          return createImagePersistenceErrorResponse(data.title, false);
        }
        const images = savedImageList(resolvedImages.images);
        const [savedAlbum] = await db
          .update(albums)
          .set({
            title: toStoredTitle(data.title),
            artist: data.artist ?? '',
            releaseDate: data.releaseDate ?? null,
            ...(data.spotifyAlbumUri
              ? { spotifyAlbumUri: data.spotifyAlbumUri }
              : {}),
          })
          .where(whereExpression)
          .returning();
        if (!savedAlbum) {
          return createMediaNotFoundEditResponse(data, type);
        }
        await replaceMediaImageRecords(
          db,
          { mediaType: 'album', mediaId: savedAlbum.id },
          resolvedImages.images,
        );

        return {
          message: `${toDisplayTitle(savedAlbum.title)} successfully edited.`,
          actionAttemptItem: toPostSavedMediaItem(savedAlbum, {
            title: toDisplayTitle(savedAlbum.title),
            images,
            spineColor: savedAlbum.spineColor,
          }),
          type,
        };
      }

      const validation = titleMediaEditSchema.safeParse(item);
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
      const storedTitle =
        type === 'movie'
          ? toStoredMovieTitle(data.title)
          : toStoredTitle(data.title);
      const displayTitle =
        type === 'movie'
          ? toDisplayMovieTitle(storedTitle)
          : toDisplayTitle(storedTitle);
      if (type === 'movie') {
        const whereExpression = data.id
          ? eq(movies.id, data.id)
          : and(
              eq(movies.title, storedTitle),
              eq(movies.movieFormat, data.movieFormat ?? 'dvd'),
            );
        const [existingMovie] = await db
          .select()
          .from(movies)
          .where(whereExpression)
          .limit(1);
        if (!existingMovie) {
          return createMediaNotFoundEditResponse(data, type);
        }
        const resolvedImages = await resolveAndPersistImageList(
          { mediaType: 'movie', mediaId: existingMovie.id },
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
        const images = savedImageList(resolvedImages.images);
        const persistedSpineColor = resolveDisplaySpineColor(
          images,
          data.spineColor,
        );
        const [savedMovie] = await db
          .update(movies)
          .set({
            title: storedTitle,
            spineColor: persistedSpineColor,
          })
          .where(whereExpression)
          .returning();
        if (!savedMovie) {
          return createMediaNotFoundEditResponse(data, type);
        }
        await replaceMediaImageRecords(
          db,
          { mediaType: 'movie', mediaId: savedMovie.id },
          resolvedImages.images,
        );

        return {
          message: `${displayTitle} successfully edited.`,
          actionAttemptItem: toPostSavedMediaItem(savedMovie, {
            title: displayTitle,
            images,
            spineColor: persistedSpineColor,
          }),
          type,
        };
      }

      const whereExpression = data.id
        ? eq(videoGames.id, data.id)
        : eq(videoGames.title, storedTitle);
      const [existingVideoGame] = await db
        .select()
        .from(videoGames)
        .where(whereExpression)
        .limit(1);
      if (!existingVideoGame) {
        return createMediaNotFoundEditResponse(data, type);
      }
      const resolvedImages = await resolveAndPersistImageList(
        { mediaType: 'videoGame', mediaId: existingVideoGame.id },
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
      const images = savedImageList(resolvedImages.images);
      const persistedSpineColor = resolveDisplaySpineColor(
        images,
        data.spineColor,
      );
      const [savedVideoGame] = await db
        .update(videoGames)
        .set({
          title: storedTitle,
          spineColor: persistedSpineColor,
        })
        .where(whereExpression)
        .returning();
      if (!savedVideoGame) {
        return createMediaNotFoundEditResponse(data, type);
      }
      await replaceMediaImageRecords(
        db,
        { mediaType: 'videoGame', mediaId: savedVideoGame.id },
        resolvedImages.images,
      );

      return {
        message: `${displayTitle} successfully edited.`,
        actionAttemptItem: toPostSavedMediaItem(savedVideoGame, {
          title: displayTitle,
          images,
          spineColor: persistedSpineColor,
        }),
        type,
      };
    }),
});
