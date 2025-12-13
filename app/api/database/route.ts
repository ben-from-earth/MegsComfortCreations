// database import
import { db } from '@/app/db/client';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import {
  MediaType,
  SuccessfulPaginationResponse,
} from '@/lib/interfaces/globalInterfaces';
import { NextRequest, NextResponse } from 'next/server';
import { albums, books, movies, videoGames } from '@/app/db/schema';
import { asc, desc, sql } from 'drizzle-orm';

const tableMap = {
  book: books,
  movie: movies,
  video_game: videoGames,
  album: albums,
};

export async function GET(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  const limit = Number(searchParams.get('limit'));
  const page = Number(searchParams.get('page'));
  const type = searchParams.get('type') as MediaType;
  const sort = searchParams.get('sort');
  const ascDesc = searchParams.get('ascDesc');

  const offset = (page - 1) * limit;
  const table = tableMap[type];
  if (!table) {
    return new ApiError(
      400,
      'Unsupported media type',
      `Unknown media type: ${type}`,
    ).format();
  }
  const sortColumn =
    type === 'book'
      ? sort === 'title'
        ? table.title
        : sort === 'pub_year'
          ? (table as typeof books).pubYear
          : table.spineColor
      : sort === 'title'
        ? table.title
        : table.spineColor;

  const orderByExpr = ascDesc === 'desc' ? desc(sortColumn) : asc(sortColumn);

  try {
    const paginatedList = await db
      .select()
      .from(table)
      .orderBy(orderByExpr)
      .limit(limit)
      .offset(offset);

    const [{ value: total }] = await db
      .select({ value: sql<number>`count(*)` })
      .from(table);

    return NextResponse.json<SuccessfulPaginationResponse>(
      {
        message: `Successful database gather`,
        paginatedList,
        total,
      },
      { status: 200 },
    );
  } catch {
    return new ApiError(
      400,
      'Database Collection Error',
      'Error gathering items from the database during pagination',
    ).format();
  }
}
