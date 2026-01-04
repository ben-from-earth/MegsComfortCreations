import { publicProcedure, router } from 'lib/trpc/trpc';
import { eq } from 'drizzle-orm';
import { nanoid, z } from 'zod';
import { db as defaultDb } from '@//db/client';
import { genres, genresBooks } from '@//db/schema';
import type {
  BookRow,
  MovieRow,
  BlockInfo,
} from 'lib/interfaces/globalInterfaces';

import { getOpenLibraryData } from './actions/get-open-library-data';
import { updateQueryCount } from 'lib/helpers/updateQueryCount';
import { getMediaCovers } from './actions/get-media-covers';
import { searchByTitle } from '../database/actions/search-by-title';

const mediaType = z.enum(['book', 'movie', 'videoGame', 'album']);

export const collectRouter = router({
  collectMedia: publicProcedure
    .input(
      z.object({
        type: mediaType,
        toCollectItem: z.object({
          title: z.string(),
          author: z.string().optional(),
        }),
      }),
    )
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const {
        type,
        toCollectItem: { title, author },
      } = input;

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
          return {
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
            blockID: nanoid(),
            isDatabase: true,
          };
        }

        const { imageUrls, title, spineColor } = mediaSearchData
          .foundMediaList[0] as MovieRow;

        //return all the block info and designate isDatabase to be true for other media
        return {
          type,
          images: imageUrls,
          blockInfo: {
            title,
            spineColor,
          },
          blockID: nanoid(),
          isDatabase: true,
        };
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
        //Just submit title as BlockInfo for non-books
        //Updates to data collection for other media types can be performed here if necessary in future update.
        blockInfo = { title };
      }

      const collectedBlockInformation = {
        type,
        images,
        blockInfo,
        blockID: nanoid(),
        isDatabase: false,
      };

      //return the collected data for creation of collectedCoverBlock
      return collectedBlockInformation;
    }),
});
