import { publicProcedure, router } from 'lib/trpc/trpc';
import { eq } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import { genres, genresBooks } from '@/db/schema';
import type {
  BookRow,
  MediaType,
  MovieRow,
  BlockInfo,
} from 'lib/interfaces/globalInterfaces';

import { getOpenLibraryData } from './actions/get-open-library-data';
import { updateQueryCount } from 'lib/helpers/updateQueryCount';
import { getMediaCovers } from './actions/get-media-covers';
import { searchByTitle } from '../database/actions/search-by-title';

// shared fields for all media
type BaseBlockInfo = {
  title: string;
  spineColor?: string;
  databaseGenres?: string[]; // only really used for books, but harmless here
};

// extra fields only for books
type BookBlockInfo = BaseBlockInfo & {
  author?: string;
  pubYear: number | null;
  pageCount: number | null;
};

// discriminated union for the whole block
export type CollectedBlockInformation =
  | {
      type: 'book';
      images: string[];
      blockInfo: BookBlockInfo;
      blockID: string;
      isDatabase: boolean;
    }
  | {
      type: 'movie' | 'videoGame' | 'album';
      images: string[];
      blockInfo: BaseBlockInfo;
      blockID: string;
      isDatabase: boolean;
    };

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

      for (const [key, searchList] of Object.entries(input)) {
        for (const item of searchList) {
          const type = key as MediaType;
          const { title, author } = item;

          const mediaSearchData = await searchByTitle(db, type, title);

          if (mediaSearchData.total > 0) {
            if (type === 'book') {
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
                type,
                images: imageUrls,
                blockInfo: {
                  title,
                  author,
                  pubYear,
                  pageCount,
                  spineColor,
                  databaseGenres,
                },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: true,
              });
            } else {
              const { imageUrls, title, spineColor } = mediaSearchData
                .foundMediaList[0] as MovieRow;

              //return all the block info and designate isDatabase to be true for other media
              blocks.push({
                type,
                images: imageUrls,
                blockInfo: {
                  title,
                  spineColor,
                },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: true,
              });
            }
          }

          //if the media wasnt in the database, collect cover images
          const images = await getMediaCovers(title, author, type);
          //conservative query count update every time we make a request to google search API.
          updateQueryCount();

          let blockInfo: BlockInfo;
          if (type === 'book') {
            //if book, go to open library and get more data about the book
            blockInfo = await getOpenLibraryData(title, author);
          } else {
            //Just submit title as blockInfo for non-books
            //Updates to data collection for other media types can be performed here if necessary in future update.
            blockInfo = { title };
          }

          const collected: CollectedBlockInformation =
            type === 'book'
              ? {
                  type: 'book',
                  images,
                  blockInfo: blockInfo as BookBlockInfo,
                  blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                  isDatabase: false,
                }
              : {
                  type,
                  images,
                  blockInfo: blockInfo as BaseBlockInfo,
                  blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                  isDatabase: false,
                };

          //return the collected data for creation of collectedCoverBlock
          blocks.push(collected);
        }
      }
      return blocks;
    }),
});
