// models
import Genre from '@/lib/database/models/genre';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import {
  PostSavedMediaItem,
  SuccessfulPaginationResponse,
} from '@/lib/interfaces/globalInterfaces';
import { NextRequest, NextResponse } from 'next/server';

export async function GET(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  //All of these options are handled by the front end so errors will be prevented before the request.
  const limit = Number(searchParams.get('limit'));
  const page = Number(searchParams.get('page'));
  const sort = searchParams.get('sort');
  const ascDesc = searchParams.get('ascDesc');

  if (!limit || !page || !sort || !ascDesc) {
    return new ApiError(
      422,
      'Missing Genre Pagination Parameters',
      'Limit, page, sort, and ascDesc are required for pagination collection',
    ).format();
  }

  const offset = (page - 1) * limit;

  try {
    const genreRes: { books: PostSavedMediaItem[]; total: number } =
      await Genre.getNoGenreBooks(sort, offset, limit, ascDesc);
    return NextResponse.json<SuccessfulPaginationResponse>(
      {
        message: `Successful database gather`,
        paginatedList: genreRes.books,
        total: genreRes.total,
      },
      { status: 200 },
    );
  } catch {
    return new ApiError(
      400,
      'Genre Error',
      'Error connecting to the database and/or genre table',
    ).format();
  }
}
