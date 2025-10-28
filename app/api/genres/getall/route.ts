// models
import Genre from '@/lib/database/models/genre';

// inerfaces and types
import { ApiError } from '@/app/api/api-Errors';
import { NextResponse } from 'next/server';

export interface getAllResponse {
  message: string;
  genres: string[];
}

export async function GET() {
  try {
    const genres: string[] = await Genre.getAllGenres();
    return NextResponse.json({
      message: 'Success',
      genres,
    });
  } catch {
    return new ApiError(
      400,
      'Genre Error',
      'Error connecting to the database and/or genre table',
    ).format();
  }
}
