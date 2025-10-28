// library imports
import axios from 'axios';

// interfaces and types
import { GoogleSearchError } from '@/app/api/api-Errors';
import { MediaType } from '@/lib/interfaces/globalInterfaces';
import { NextRequest, NextResponse } from 'next/server';

export interface GoogleSearchResponse {
  items?: { link: string }[];
}

const API_KEY = process.env.GOOGLE_SEARCH_API_KEY;
const CX = process.env.GOOGLE_SEARCH_CX;

export async function POST(req: NextRequest) {
  const body: { title: string; author?: string; type: MediaType } =
    await req.json();
  const { title, author, type } = body;

  const imgArr: string[] = [];
  try {
    if (!CX || !API_KEY) {
      throw new GoogleSearchError(
        401,
        'Google Search Credential Error',
        'Error Connecting to Google Search API because of invalid or empty credentials',
      );
    }
    const params = new URLSearchParams({
      q: `${title} ${author ? ` ${author}` : ''} ${type} Cover Image`,
      cx: CX,
      key: API_KEY,
      searchType: 'image',
      num: '3',
    });
    const { data } = await axios.get<GoogleSearchResponse>(
      `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
    );

    const items = data.items ?? [];
    items.map((i) => imgArr.push(i.link));
    return NextResponse.json({ images: imgArr }, { status: 200 });
  } catch (error) {
    if (error instanceof GoogleSearchError) {
      return error.format();
    } else {
      return new GoogleSearchError(
        400,
        'Google Search Error',
        'Error Connecting to Google Search API',
      ).format();
    }
  }
}
