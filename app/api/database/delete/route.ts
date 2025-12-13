import { db } from '@/app/db/client';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import { NextRequest, NextResponse } from 'next/server';
import { ilike } from 'drizzle-orm';
import { books, movies, videoGames, albums } from '@/app/db/schema';

const tableMap = {
  book: books,
  movie: movies,
  video_game: videoGames,
  album: albums,
} as const;

type DeletableType = keyof typeof tableMap;

export async function DELETE(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  const type = searchParams.get('type') as DeletableType | null;
  const title = searchParams.get('title');

  if (!type || !title) {
    return new ApiError(
      400,
      'Bad Request',
      '`type` and `title` query parameters are required.',
    ).format();
  }

  const table = tableMap[type];

  if (!table) {
    return new ApiError(
      400,
      'Bad Request',
      `Unsupported type: ${type}.`,
    ).format();
  }

  try {
    // Drizzle's delete() returns an array of deleted rows (if .returning() is used).
    const deleted = await db
      .delete(table)
      .where(ilike(table.title, title)) // same as "WHERE title ILIKE $1"
      .returning({ id: table.id });

    if (deleted.length === 0) {
      throw new ApiError(
        404,
        'Non-existant Deletion Error',
        `No item with title: ${title} in the ${type} database exists`,
      );
    }

    return NextResponse.json(
      {
        message: `Successfully deleted ${title}`,
      },
      { status: 200 },
    );
  } catch (error) {
    if (error instanceof ApiError) {
      return error.format();
    } else {
      return new ApiError(
        400,
        'Database Deletion Error',
        'Error deleting items from the database.',
      ).format();
    }
  }
}
