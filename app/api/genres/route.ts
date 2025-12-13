// models
import Genre from '@/lib/database/models/genre';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import { NextRequest, NextResponse } from 'next/server';
import { BookRow } from '@/lib/interfaces/globalInterfaces';

export async function GET(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  //All of these options are handled by the front end so errors will be prevented before the request.
  const genre = searchParams.get('genre');
  const limit = Number(searchParams.get('limit'));
  const page = Number(searchParams.get('page'));
  const sort = searchParams.get('sort');
  const ascDesc = searchParams.get('ascDesc');

  if (!genre || !limit || !page || !sort || !ascDesc) {
    return new ApiError(
      422,
      'Missing Genre Pagination Parameters',
      'Genre, limit, page, sort, and ascDesc are required for pagination collection',
    ).format();
  }

  const offset = (page - 1) * limit;
  try {
    const genreRes: { books: BookRow[]; total: number } =
      await Genre.getBooksWithGenre(genre, sort, offset, limit, ascDesc);
    return NextResponse.json(
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
