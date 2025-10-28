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
import { NextRequest, NextResponse } from 'next/server';
import {
  MediaType,
  postSavedMediaItem,
  presavedMediaItem,
  SuccessfulMediaSaveEditResponse,
} from '@/lib/interfaces/globalInterfaces';
import {
  ApiError,
  DatabaseSaveEditErrorResponse,
  ErrorResponse,
  PGDatabaseError,
  SchemaViolationError,
} from '@/app/api/api-Errors';

import { DatabaseError } from 'pg';

export async function POST(
  req: NextRequest,
  { params }: { params: { type: MediaType } },
): Promise<
  | NextResponse<SuccessfulMediaSaveEditResponse>
  | NextResponse<DatabaseSaveEditErrorResponse>
  | NextResponse<ErrorResponse>
> {
  const { type } = await params;
  const body: presavedMediaItem = await req.json();

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
        const book: postSavedMediaItem = await Book.create(body);

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: book,
            type,
          },
          { status: 201 },
        );
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else if (error instanceof DatabaseError) {
          if (error.code === '23505') {
            return new PGDatabaseError(body, type, error.detail!).format();
          }
        } else {
          return new ApiError(
            400,
            'Database Error',
            'Database error during save attempt',
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
        const movie: postSavedMediaItem = await Movie.create(body);

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: movie,
            type,
          },
          { status: 201 },
        );
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else if (error instanceof DatabaseError) {
          if (error.code === '23505') {
            return new PGDatabaseError(body, type, error.detail!).format();
          }
        } else {
          return new ApiError(
            400,
            'Database Error',
            'Database error during save attempt',
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
        const video_game: postSavedMediaItem = await Video_Game.create(body);

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: video_game,
            type,
          },
          { status: 201 },
        );
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else if (error instanceof DatabaseError) {
          if (error.code === '23505') {
            return new PGDatabaseError(body, type, error.detail!).format();
          }
        } else {
          return new ApiError(
            400,
            'Database Error',
            'Database error during save attempt',
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
        const album: postSavedMediaItem = await Album.create(body);

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: album,
            type,
          },
          { status: 201 },
        );
      } catch (error) {
        if (error instanceof SchemaViolationError) {
          return error.format();
        } else if (error instanceof DatabaseError) {
          if (error.code === '23505') {
            return new PGDatabaseError(body, type, error.detail!).format();
          }
        } else {
          return new ApiError(
            400,
            'Database Error',
            'Database error during save attempt',
          ).format();
        }
      }
    default:
      // handle unsupported type
      return new ApiError(
        400,
        'Unsupported type',
        `Unknown save type: ${type}`,
      ).format();
  }
}
