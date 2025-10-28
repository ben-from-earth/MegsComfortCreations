// database import
import db from '@/lib/database/db';

// interfaces and types
import { ApiError } from '@/app/api/api-Errors';
import { NextRequest, NextResponse } from 'next/server';

export async function DELETE(req: NextRequest) {
  const { searchParams } = req.nextUrl;

  const type = searchParams.get('type');
  const title = searchParams.get('title');

  try {
    const deleteRes = await db.query(
      `DELETE FROM ${type + 's'}
          WHERE title ILIKE $1`,
      [title],
    );
    if (deleteRes.rowCount === 0) {
      throw new ApiError(
        404,
        'Non-existant Deletion Error',
        `No item with title:${title} in the ${type} database exists`,
      );
    } else {
      return NextResponse.json(
        {
          message: `Successfully deleted ${title}`,
        },
        { status: 200 },
      );
    }
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
