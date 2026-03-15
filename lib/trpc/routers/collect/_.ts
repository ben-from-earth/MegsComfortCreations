import { adminProcedure, router } from 'lib/trpc/trpc';
import { eq } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import { genres, genresBooks } from '@/db/schema';
import type { MediaType } from 'lib/interfaces/globalInterfaces';

import { getOpenLibraryData } from './actions/get-open-library-data';
import { getMediaCovers } from './actions/get-media-covers';
import { searchByTitle } from '../database/actions/search-by-title';
import { googleApiQueryUsage } from '@/db/schema';
import { collectedBlockInformationSchema } from '@/mediacollector/collector-form/collectorFormSchema';
export type CollectedBlockInformation = z.infer<
  typeof collectedBlockInformationSchema
>;

export const collectRouter = router({
  collectMedia: adminProcedure
    .input(
      z.object({
        book: z.array(
          z.object({
            title: z.string(),
            author: z.string().optional(),
          }),
        ),
        movie: z.array(
          z.object({
            title: z.string(),
            author: z.string().optional(),
          }),
        ),
        videoGame: z.array(
          z.object({
            title: z.string(),
            author: z.string().optional(),
          }),
        ),
        album: z.array(
          z.object({
            title: z.string(),
            author: z.string().optional(),
          }),
        ),
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
            const foundMedia = mediaSearchData.foundMediaList[0];
            if (!foundMedia) {
              continue;
            }
            if (type === 'book') {
              if (!('author' in foundMedia)) {
                continue;
              }
              const {
                id,
                imageUrls,
                title: foundTitle,
                author,
                pageCount,
                pubYear,
                spineColor,
              } = foundMedia;
              const rows = await db
                .select({ genre: genres.genre })
                .from(genres)
                .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
                .where(eq(genresBooks.bookId, id));

              const databaseGenres = rows.map((row) => row.genre);

              blocks.push({
                type: 'book',
                images: imageUrls.map((url) => ({ url, selected: false })),
                blockInfo: {
                  title: foundTitle,
                  author,
                  pubYear,
                  pageCount,
                  spineColor,
                  genres: databaseGenres,
                },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: true,
              });
            } else {
              const { imageUrls, title: foundTitle, spineColor } = foundMedia;

              blocks.push({
                type,
                images: imageUrls.map((url) => ({ url, selected: false })),
                blockInfo: {
                  title: foundTitle,
                  spineColor,
                  genres: [],
                },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: true,
              });
            }
          } else {
            //if the media wasnt in the database, collect cover images
            const images = await getMediaCovers(title, author, type);
            //conservative query count update every time we make a request to google search API.
            todayQueryCount += 1;
            if (type === 'book') {
              const bookInfo = await getOpenLibraryData(title, author);
              blocks.push({
                type,
                images: images.map((url) => ({ url, selected: false })),
                blockInfo: { ...bookInfo, spineColor: '#ffffff', genres: [] },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: false,
              });
            } else {
              blocks.push({
                type,
                images: images.map((url) => ({ url, selected: false })),
                blockInfo: { title, spineColor: '#ffffff', genres: [] },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: false,
              });
            }
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
