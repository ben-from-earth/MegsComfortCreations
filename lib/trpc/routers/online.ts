import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import axios from 'axios';

const API_KEY = process.env.GOOGLE_SEARCH_API_KEY;
const CX = process.env.GOOGLE_SEARCH_CX;

export interface OpenLibrarySuccess {
  title: string;
  author: string;
  pubYear: number;
  pageCount: number;
}

export const onlineRouter = router({
  openLibrary: adminProcedure
    .input(
      z.object({
        title: z.string().min(1),
        author: z.string().min(1).optional(),
      }),
    )
    .mutation(async ({ input }) => {
      if (!input.author) {
        return {
          error: 'Open Library Error',
          message: `Error gathering Open Library data for ${input.title}, author not provided`,
          failedSearchData: { title: input.title, author: input.author },
        };
      }
      const params = new URLSearchParams({
        title: input.title,
        author: input.author,
        limit: '1',
        fields: 'first_publish_year,number_of_pages_median',
      });
      const openLibraryRes = await axios.get(
        `https://openlibrary.org/search.json?${params.toString()}`,
      );
      const doc = openLibraryRes.data?.docs?.[0];
      if (!doc) {
        return {
          error: 'Open Library Error',
          message: `Error gathering Open Library data for ${input.title}`,
          failedSearchData: { title: input.title, author: input.author },
        };
      }
      const {
        first_publish_year: pubYear,
        number_of_pages_median: pageCount,
      }: { first_publish_year: number; number_of_pages_median: number } = doc;
      return { title: input.title, author: input.author, pubYear, pageCount };
    }),

  mediaCovers: adminProcedure
    .input(
      z.object({
        title: z.string().min(1),
        author: z.string().optional(),
        type: z.enum(['book', 'movie', 'videoGame', 'album']),
      }),
    )
    .mutation(async ({ input }) => {
      const imgArr: string[] = [];
      if (!CX || !API_KEY) {
        return {
          error: 'Google Search Credential Error',
          message:
            'Error Connecting to Google Search API because of invalid or empty credentials',
          failedSearchData: [],
        };
      }
      const params = new URLSearchParams({
        q: `${input.title}${input.author ? ` ${input.author}` : ''} ${input.type} Cover Image`,
        cx: CX,
        key: API_KEY,
        searchType: 'image',
        num: '3',
      });
      const { data } = await axios.get<{ items?: { link: string }[] }>(
        `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
      );
      (data.items ?? []).forEach((i) => imgArr.push(i.link));
      return { images: imgArr };
    }),
});
