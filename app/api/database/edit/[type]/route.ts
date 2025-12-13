// library imports
import { validate } from 'jsonschema';

// schemas
import bookCreateSchema from '@/lib/database/schemas/bookCreateSchema.json';
import otherMediaCreateSchema from '@/lib/database/schemas/otherMediaCreateSchema.json';

// drizzle
import { db } from '@/app/db/client'; // adjust to your actual db path
import { books, movies, videoGames, albums } from '@/app/db/schema';
import { eq } from 'drizzle-orm';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// interfaces and types
import {
  ApiError,
  DatabaseSaveEditErrorResponse,
  SchemaViolationError,
} from '@/app/api/api-Errors';
import {
  AlbumRow,
  BookRow,
  MovieRow,
  PostSavedMediaItem,
  SuccessfulMediaSaveEditResponse,
  VideoGameRow,
} from '@/lib/interfaces/globalInterfaces';
import { NextRequest, NextResponse } from 'next/server';

export async function PUT(
  req: NextRequest,
  { params }: { params: Promise<{ type: string }> },
) {
  const { type } = await params;
  const body: PostSavedMediaItem = await req.json();

  switch (type) {
    case 'book':
      try {
        const validation = validate(body, bookCreateSchema);
        if (!validation.valid) {
          throw new SchemaViolationError(
            validation.errors.map((e) => e.stack),
            body,
            type,
          );
        }

        // Prefer id if present, otherwise fall back to title-based match.
        const whereExpr = body.id
          ? eq(books.id, body.id as string)
          : eq(books.title, body.title);

        const [book] = await db
          .update(books)
          .set({
            // Adjust these to your actual body keys as needed:
            title: (body as BookRow).title,
            author: (body as BookRow).author,
            pageCount: (body as BookRow).pageCount,
            pubYear: (body as BookRow).pubYear,
            spineColor: (body as BookRow).spineColor,
            imageUrls: (body as BookRow).imageUrls,
          })
          .where(whereExpr)
          .returning();

        if (!book) {
          return NextResponse.json<DatabaseSaveEditErrorResponse>(
            {
              error: 'Media Not Found',
              message:
                'Edit requested on an item that does not exist in the database',
              actionAttemptItem: body,
              type,
              errors: [`${body.title} does not exist in the database.`],
            },
            { status: 404 },
          );
        } else {
          return NextResponse.json<SuccessfulMediaSaveEditResponse>(
            {
              message: `${titleRearrange(book.title)} successfully edited.`,
              actionAttemptItem: book,
              type,
            },
            { status: 200 },
          );
        }
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else {
          return new ApiError(
            400,
            'Edit Error',
            `Edit request of ${body.title} failed`,
          ).format();
        }
      }

    case 'movie':
      try {
        const validation = validate(body, otherMediaCreateSchema);
        if (!validation.valid) {
          throw new SchemaViolationError(
            validation.errors.map((e) => e.stack),
            body,
            type,
          );
        }

        const whereExpr = body.id
          ? eq(movies.id, body.id as string)
          : eq(movies.title, body.title);

        const [movie] = await db
          .update(movies)
          .set({
            title: body.title,
            spineColor: (body as MovieRow).spineColor,
            imageUrls: (body as MovieRow).imageUrls,
          })
          .where(whereExpr)
          .returning();

        if (!movie) {
          return NextResponse.json<DatabaseSaveEditErrorResponse>(
            {
              error: 'Media Not Found',
              message:
                'Edit requested on an item that does not exist in the database',
              actionAttemptItem: body,
              type,
              errors: [`${body.title} does not exist in the database.`],
            },
            { status: 404 },
          );
        } else {
          return NextResponse.json<SuccessfulMediaSaveEditResponse>(
            {
              message: `${titleRearrange(movie.title)} successfully edited.`,
              actionAttemptItem: movie,
              type,
            },
            { status: 200 },
          );
        }
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else {
          return new ApiError(
            400,
            'Edit Error',
            `Edit request of ${body.title} failed`,
          ).format();
        }
      }

    case 'video_game':
      try {
        const validation = validate(body, otherMediaCreateSchema);
        if (!validation.valid) {
          throw new SchemaViolationError(
            validation.errors.map((e) => e.stack),
            body,
            type,
          );
        }

        const whereExpr = body.id
          ? eq(videoGames.id, body.id as string)
          : eq(videoGames.title, body.title);

        const [videoGame] = await db
          .update(videoGames)
          .set({
            title: body.title,
            spineColor: (body as VideoGameRow).spineColor,
            imageUrls: (body as VideoGameRow).imageUrls,
          })
          .where(whereExpr)
          .returning();

        if (!videoGame) {
          return NextResponse.json<DatabaseSaveEditErrorResponse>(
            {
              error: 'Media Not Found',
              message:
                'Edit requested on an item that does not exist in the database',
              actionAttemptItem: body,
              type,
              errors: [`${body.title} does not exist in the database.`],
            },
            { status: 404 },
          );
        } else {
          return NextResponse.json<SuccessfulMediaSaveEditResponse>(
            {
              message: `${titleRearrange(videoGame.title)} successfully edited.`,
              actionAttemptItem: videoGame,
              type,
            },
            { status: 200 },
          );
        }
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else {
          return new ApiError(
            400,
            'Edit Error',
            `Edit request of ${body.title} failed`,
          ).format();
        }
      }

    case 'album':
      try {
        const validation = validate(body, otherMediaCreateSchema);
        if (!validation.valid) {
          throw new SchemaViolationError(
            validation.errors.map((e) => e.stack),
            body,
            type,
          );
        }

        const whereExpr = body.id
          ? eq(albums.id, body.id as string)
          : eq(albums.title, body.title);

        const [album] = await db
          .update(albums)
          .set({
            title: body.title,
            spineColor: (body as AlbumRow).spineColor,
            imageUrls: (body as AlbumRow).imageUrls,
          })
          .where(whereExpr)
          .returning();

        if (!album) {
          return NextResponse.json<DatabaseSaveEditErrorResponse>(
            {
              error: 'Media Not Found',
              message:
                'Edit requested on an item that does not exist in the database',
              actionAttemptItem: body,
              type,
              errors: [`${body.title} does not exist in the database.`],
            },
            { status: 404 },
          );
        } else {
          return NextResponse.json<SuccessfulMediaSaveEditResponse>(
            {
              message: `${titleRearrange(album.title)} successfully edited.`,
              actionAttemptItem: album,
              type,
            },
            { status: 200 },
          );
        }
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else {
          return new ApiError(
            400,
            'Edit Error',
            `Edit request of ${body.title} failed`,
          ).format();
        }
      }

    default:
      return new ApiError(
        400,
        'Bad Request',
        `Unsupported media type: ${type}`,
      ).format();
  }
}
