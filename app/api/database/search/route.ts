// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// drizzle
import { db } from '@/app/db/client';
import { books, movies, videoGames, albums } from '@/app/db/schema';
import { ilike } from 'drizzle-orm';

// interfaces and types
import { NextRequest, NextResponse } from 'next/server';
import { MediaType } from '@/lib/interfaces/globalInterfaces';
import { ApiError } from '@/app/api/api-Errors';

export async function GET(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  const typeParam = searchParams.get('type');
  const titleParam = searchParams.get('title');

  if (!typeParam || !titleParam) {
    return new ApiError(
      422,
      'Missing Search Parameters',
      'Search parameters `type` and `title` are required',
    ).format();
  }

  const type = typeParam as MediaType;
  const title: string = titleRearrange(titleParam);

  switch (type) {
    case 'book':
      try {
        const result = await db
          .select()
          .from(books)
          .where(ilike(books.title, title));
        const total = result.length;

        if (total === 0) {
          throw new ApiError(
            404,
            'Media not found',
            `No ${type} in database with title ${title}`,
          );
        }

        return NextResponse.json(
          {
            message: `Successfully found ${total} ${type}(s) with title ${titleRearrange(
              result[0].title,
            )}`,
            foundMediaList: result,
            total,
          },
          { status: 200 },
        );
      } catch (error) {
        if (error instanceof ApiError) {
          return error.format();
        }
        return new ApiError(
          400,
          'Database collection error',
          'Error gathering items from the database during search',
        ).format();
      }

    case 'movie':
      try {
        const result = await db
          .select()
          .from(movies)
          .where(ilike(movies.title, title));
        const total = result.length;

        if (total === 0) {
          throw new ApiError(
            404,
            'Media not found',
            `No ${type} in database with title ${title}`,
          );
        }

        return NextResponse.json(
          {
            message: `Successfully found ${total} ${type}(s) with title ${titleRearrange(
              result[0].title,
            )}`,
            foundMediaList: result,
            total,
          },
          { status: 200 },
        );
      } catch (error) {
        if (error instanceof ApiError) {
          return error.format();
        }
        return new ApiError(
          400,
          'Database collection error',
          'Error gathering items from the database during search',
        ).format();
      }

    case 'video_game':
      try {
        const result = await db
          .select()
          .from(videoGames)
          .where(ilike(videoGames.title, title));
        const total = result.length;

        if (total === 0) {
          throw new ApiError(
            404,
            'Media not found',
            `No ${type} in database with title ${title}`,
          );
        }

        return NextResponse.json(
          {
            message: `Successfully found ${total} ${type}(s) with title ${titleRearrange(
              result[0].title,
            )}`,
            foundMediaList: result,
            total,
          },
          { status: 200 },
        );
      } catch (error) {
        if (error instanceof ApiError) {
          return error.format();
        }
        return new ApiError(
          400,
          'Database collection error',
          'Error gathering items from the database during search',
        ).format();
      }

    case 'album':
      try {
        const result = await db
          .select()
          .from(albums)
          .where(ilike(albums.title, title));
        const total = result.length;

        if (total === 0) {
          throw new ApiError(
            404,
            'Media not found',
            `No ${type} in database with title ${title}`,
          );
        }

        return NextResponse.json(
          {
            message: `Successfully found ${total} ${type}(s) with title ${titleRearrange(
              result[0].title,
            )}`,
            foundMediaList: result,
            total,
          },
          { status: 200 },
        );
      } catch (error) {
        if (error instanceof ApiError) {
          return error.format();
        }
        return new ApiError(
          400,
          'Database collection error',
          'Error gathering items from the database during search',
        ).format();
      }

    default:
      return new ApiError(
        400,
        'Unsupported type',
        `Unknown media type: ${type}`,
      ).format();
  }
}
