// database import
import db from '@/lib/database/db';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import {
  postSavedMediaItem,
  SuccessfulPaginationResponse,
} from '@/lib/interfaces/globalInterfaces';
import { NextRequest, NextResponse } from 'next/server';

export async function GET(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  const limit = Number(searchParams.get('limit'));
  const page = Number(searchParams.get('page'));
  const type = searchParams.get('type');
  const sort = searchParams.get('sort');
  const ascDesc = searchParams.get('ascDesc');

  const offset = (page - 1) * limit;

  try {
    const result = await db.query<postSavedMediaItem>(
      `SELECT * 
          FROM ${type + 's'}
          ORDER BY ${sort} ${ascDesc}
          LIMIT $1 OFFSET $2`,
      [limit, offset],
    );
    const paginatedList = result.rows;

    const totalRes = await db.query(
      `SELECT COUNT(*) 
          FROM ${type + 's'}`,
    );

    const total = parseInt(totalRes.rows[0].count, 10);
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
