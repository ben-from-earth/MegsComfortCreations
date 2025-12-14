// library imports
import { validate } from 'jsonschema';
import { DatabaseError } from 'pg';

// schemas
import bookCreateSchema from '@/lib/database/schemas/bookCreateSchema.json';
import otherMediaCreateSchema from '@/lib/database/schemas/otherMediaCreateSchema.json';

// drizzle
import { db } from '@/app/db/client'; // adjust path to your drizzle instance
import { books, movies, videoGames, albums } from '@/app/db/schema';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// interfaces and types
import { NextRequest, NextResponse } from 'next/server';
import {
  BookInsert,
  PreSavedMediaItem,
  SuccessfulMediaSaveEditResponse,
} from '@/lib/interfaces/globalInterfaces';
import {
  ApiError,
  DatabaseSaveEditErrorResponse,
  ErrorResponse,
  PGDatabaseError,
  SchemaViolationError,
} from '@/app/api/api-Errors';

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ type: string }> },
): Promise<
  | NextResponse<SuccessfulMediaSaveEditResponse>
  | NextResponse<DatabaseSaveEditErrorResponse>
  | NextResponse<ErrorResponse>
> {
  const { type } = await params;
  const body: PreSavedMediaItem = await req.json();

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

        const [book] = await db
          .insert(books)
          .values({
            title: (body as BookInsert).title,
            author: (body as BookInsert).author,
            pageCount: (body as BookInsert).pageCount ?? null,
            pubYear: (body as BookInsert).pubYear ?? null,
            spineColor: (body as BookInsert).spineColor,
            imageUrls: (body as BookInsert).imageUrls,
          })
          .returning();

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: {
              ...book,
              genres: body.genres,
              blockID: body.blockID,
            },
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
        }
        return new ApiError(
          400,
          'Database Error',
          'Database error during save attempt',
        ).format();
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

        const [movie] = await db
          .insert(movies)
          .values({
            title: body.title,
            spineColor: body.spineColor,
            imageUrls: body.imageUrls,
          })
          .returning();

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: {
              ...movie,
              blockID: body.blockID,
            },
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
        }
        return new ApiError(
          400,
          'Database Error',
          'Database error during save attempt',
        ).format();
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

        const [videoGame] = await db
          .insert(videoGames)
          .values({
            title: body.title,
            spineColor: body.spineColor,
            imageUrls: body.imageUrls,
          })
          .returning();

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: {
              ...videoGame,
              blockID: body.blockID,
            },
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
        }
        return new ApiError(
          400,
          'Database Error',
          'Database error during save attempt',
        ).format();
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

        const [album] = await db
          .insert(albums)
          .values({
            title: body.title,
            spineColor: body.spineColor,
            imageUrls: body.imageUrls,
          })
          .returning();

        return NextResponse.json<SuccessfulMediaSaveEditResponse>(
          {
            message: `${titleRearrange(
              body.title,
            )} successfully added to database.`,
            actionAttemptItem: {
              ...album,
              blockID: body.blockID,
            },
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
        }
        return new ApiError(
          400,
          'Database Error',
          'Database error during save attempt',
        ).format();
      }

    default:
      return new ApiError(
        400,
        'Unsupported type',
        `Unknown save type: ${type}`,
      ).format();
  }
}
