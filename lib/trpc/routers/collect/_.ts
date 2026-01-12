import { publicProcedure, router } from 'lib/trpc/trpc';
import { eq } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import { genres, genresBooks } from '@/db/schema';
import type { BookRow, MediaType } from 'lib/interfaces/globalInterfaces';

import { getOpenLibraryData } from './actions/get-open-library-data';
import { getMediaCovers } from './actions/get-media-covers';
import { searchByTitle } from '../database/actions/search-by-title';
import { googleApiQueryUsage } from '@/db/schema';

// shared fields for all media
type BaseBlockInfo = {
  title: string;
  spineColor: string;
  genres: string[];
};

// extra fields only for books
type BookBlockInfo = BaseBlockInfo & {
  author: string | null;
  pubYear: number | null;
  pageCount: number | null;
};

// discriminated union for the whole block
export type CollectedBlockInformation = {
  type: MediaType;
  images: { url: string; selected: boolean }[];
  blockInfo: BookBlockInfo;
  blockID: string;
  isDatabase: boolean;
};
// | {
//     type: 'movie' | 'videoGame' | 'album';
//     images: string[];
//     blockInfo: BaseBlockInfo;
//     blockID: string;
//     isDatabase: boolean;
//   };

export const collectRouter = router({
  collectMedia: publicProcedure
    .input(
      z.object({
        book: z.array(
          z.object({
            title: z.string(),
            author: z.string().optional(),
          }),
        ),
        // movie: z.array(
        //   z.object({
        //     title: z.string(),
        //     author: z.string().optional(),
        //   }),
        // ),
        // videoGame: z.array(
        //   z.object({
        //     title: z.string(),
        //     author: z.string().optional(),
        //   }),
        // ),
        // album: z.array(
        //   z.object({
        //     title: z.string(),
        //     author: z.string().optional(),
        //   }),
        // ),
      }),
    )
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const blocks: CollectedBlockInformation[] = [];

      const todayStr = new Date().toLocaleDateString('en-CA', {
        timeZone: 'America/New_York',
      });

      const usageRows = await db
        .select({ queryCount: googleApiQueryUsage.queryCount })
        .from(googleApiQueryUsage)
        .where(eq(googleApiQueryUsage.date, todayStr))
        .limit(1);

      let todayQueryCount = usageRows[0]?.queryCount ?? 0;

      for (const [key, searchList] of Object.entries(input)) {
        const type = key as MediaType;
        for (const item of searchList) {
          const { title, author } = item;

          const mediaSearchData = await searchByTitle(db, type, title);

          if (mediaSearchData.total > 0) {
            const {
              id,
              imageUrls,
              title,
              author,
              pageCount,
              pubYear,
              spineColor,
            } = mediaSearchData.foundMediaList[0] as BookRow;
            //get genres tied to the found book id
            const rows = await db
              .select({ genre: genres.genre })
              .from(genres)
              .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
              .where(eq(genresBooks.bookId, id));

            const databaseGenres = rows.map((row) => row.genre);

            //return all the block info and designate isDatabase to be true for books
            blocks.push({
              type: 'book',
              images: imageUrls.map((url) => ({ url, selected: false })),
              blockInfo: {
                title,
                author,
                pubYear,
                pageCount,
                spineColor,
                genres: databaseGenres,
              },
              blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
              isDatabase: true,
            });

            // else {
            //   const { imageUrls, title, spineColor } = mediaSearchData
            //     .foundMediaList[0] as MovieRow;

            //   //return all the block info and designate isDatabase to be true for other media
            //   blocks.push({
            //     type,
            //     images: imageUrls,
            //     blockInfo: {
            //       title,
            //       spineColor,
            //     },
            //     blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
            //     isDatabase: true,
            //   });
            // }
          } else {
            //if the media wasnt in the database, collect cover images
            const images = await getMediaCovers(title, author, type);
            //conservative query count update every time we make a request to google search API.
            todayQueryCount += 1;
            // if (type === 'book') {
            //if book, go to open library and get more data about the book
            const blockInfo = await getOpenLibraryData(title, author);
            // } else {
            //   //Just submit title as blockInfo for non-books
            //   //Updates to data collection for other media types can be performed here if necessary in future update.
            //   blockInfo = { title };
            // }

            const collected = {
              type,
              images: images.map((url) => ({ url, selected: false })),
              blockInfo: { ...blockInfo, spineColor: '#ffffff', genres: [] },
              blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
              isDatabase: false,
            };
            // : {
            //     type,
            //     images,
            //     blockInfo: blockInfo as BaseBlockInfo,
            //     blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
            //     isDatabase: false,
            //   };

            //return the collected data for creation of collectedCoverBlock
            blocks.push(collected);
          }
        }
      }
      //update today's google api query count in the database
      await db
        .update(googleApiQueryUsage)
        .set({ queryCount: todayQueryCount })
        .where(eq(googleApiQueryUsage.date, todayStr));
      const databaseSorted = [...blocks].sort((a, b) =>
        a.isDatabase === b.isDatabase ? 0 : a.isDatabase ? 1 : -1,
      );
      return databaseSorted;
    }),
});
