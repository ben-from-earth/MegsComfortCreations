// models
import Genre from '@/lib/database/models/genre';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import { NextRequest, NextResponse } from 'next/server';

export async function GET(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  const bookID = searchParams.get('bookID');
  try {
    if (!bookID) {
      throw new ApiError(422, 'Genre Error', 'BookID not provided');
    }
    const genres = await Genre.getforbook(bookID);
    return NextResponse.json(
      {
        message: `Successfully grabbed genres for bookID ${bookID}`,
        genres,
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
