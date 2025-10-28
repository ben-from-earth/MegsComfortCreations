// library imports
import axios from 'axios';

// interfaces and types
import { OpenLibraryError } from '@/app/api/api-Errors';
import { NextRequest, NextResponse } from 'next/server';

export interface OpenLibrarySuccess {
  title: string;
  author: string;
  pub_year: number;
  page_count: number;
}

export async function POST(req: NextRequest) {
  const body: { title: string; author: string } = await req.json();
  const { title, author } = body;
  try {
    const params = new URLSearchParams({
      title,
      author,
      limit: '1',
      fields: 'first_publish_year,number_of_pages_median',
    });
    const openLibraryRes = await axios.get(
      `https://openlibrary.org/search.json?${params.toString()}`,
    );

    const doc = openLibraryRes.data?.docs?.[0];
    if (!doc) {
      return new OpenLibraryError(
        400,
        'Open Library Error',
        `Error gathering Open Library data for ${title}`,
        { title, author },
      ).format();
    } else {
      const {
        first_publish_year: pub_year,
        number_of_pages_median: page_count,
      }: { first_publish_year: number; number_of_pages_median: number } = doc;
      return NextResponse.json<OpenLibrarySuccess>(
        {
          title,
          author,
          pub_year,
          page_count,
        },
        { status: 200 },
      );
    }
  } catch {
    return new OpenLibraryError(
      400,
      'Open Library Error',
      `Error gathering Open Library data for ${title}`,
      { title, author },
    ).format();
  }
}
