// library imports
import { validate } from 'jsonschema';

// schemas
import bookCreateSchema from '@/lib/database/schemas/bookCreateSchema.json';
import otherMediaCreateSchema from '@/lib/database/schemas/otherMediaCreateSchema.json';

// models
import Book from '@/lib/database/models/book';
import Movie from '@/lib/database/models/movie';
import Video_Game from '@/lib/database/models/video_game';
import Album from '@/lib/database/models/album';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// interfaces and types
import {
  ApiError,
  DatabaseSaveEditErrorResponse,
  SchemaViolationError,
} from '@/app/api/api-Errors';
import {
  MediaType,
  postSavedMediaItem,
  SuccessfulMediaSaveEditResponse,
} from '@/lib/interfaces/globalInterfaces';
import { NextRequest, NextResponse } from 'next/server';

export async function PUT(
  req: NextRequest,
  { params }: { params: Promise<{ type: MediaType }> },
) {
  const { type } = await params;
  const body: postSavedMediaItem = await req.json();
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
        const book = await Book.edit(body);
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
              type: type,
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
        const movie = await Movie.edit(body);
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
              type: type,
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
        const video_game = await Video_Game.edit(body);
        if (!video_game) {
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
              message: `${titleRearrange(video_game.title)} successfully edited.`,
              actionAttemptItem: video_game,
              type: type,
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
        const album = await Album.edit(body);
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
              type: type,
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
  }
}
