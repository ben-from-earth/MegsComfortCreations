import { adminProcedure, router } from 'lib/trpc/trpc';
import { TRPCError } from '@trpc/server';
import { eq, sql } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb } from '@/db/client';
import { genres, genresBooks } from '@/db/schema';
import { MediaType } from 'lib/constants/mediaTypes';

import { getOpenLibraryData } from './actions/get-open-library-data';
import { getMediaCovers } from './actions/get-media-covers';
import { searchByTitle } from '../database/actions/search-by-title';
import { googleApiQueryUsage } from '@/db/schema';
import type { CollectedBlockInformation } from '@/mediacollector/collector-form/collectorFormSchema';
import { persistUploadedImageToS3 } from 'lib/media-storage/local-image-storage';

export const collectRouter = router({
  uploadCoverImage: adminProcedure
    .input(
      z.object({
        blockID: z.string().min(1),
        sortOrder: z.number().int().nonnegative(),
        fileName: z.string().min(1),
        mimeType: z.string().min(1),
        dataBase64: z.string().min(1),
      }),
    )
    .mutation(async ({ input }) => {
      const fileBuffer = Buffer.from(input.dataBase64, 'base64');
      if (!fileBuffer || fileBuffer.length === 0) {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: 'Uploaded file payload is empty or invalid.',
        });
      }

      const uploadedImage = await persistUploadedImageToS3({
        imageBuffer: fileBuffer,
        mediaType: 'book',
        mediaId: input.blockID,
        sortOrder: input.sortOrder,
        fileName: input.fileName,
        mimeType: input.mimeType,
      });

      return {
        url: uploadedImage.publicPath,
        isDefault: input.sortOrder === 0,
        spineColor: '#ffffff',
        selected: true,
      };
    }),
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

      let googleApiQueryIncrement = 0;

      for (const [key, searchList] of Object.entries(input)) {
        const type = key as MediaType;
        for (const item of searchList) {
          const { title, author } = item;

          const mediaSearchData = await searchByTitle(db, type, title, author);

          if (mediaSearchData.total > 0) {
            for (const foundMedia of mediaSearchData.foundMediaList) {
              if (type === 'book') {
                if (!('author' in foundMedia)) {
                  continue;
                }
                const {
                  id,
                  title: foundTitle,
                  author: foundAuthor,
                  pageCount,
                  pubYear,
                  spineColor,
                } = foundMedia;
                const normalizedImages =
                  'images' in foundMedia && Array.isArray(foundMedia.images)
                    ? foundMedia.images
                    : ('imageUrls' in foundMedia && Array.isArray(foundMedia.imageUrls)
                        ? foundMedia.imageUrls.map((url) => ({
                            url,
                            isDefault: false,
                            spineColor,
                            selected: false,
                          }))
                        : []);
                if (normalizedImages.length > 0 && !normalizedImages.some((image) => image.isDefault)) {
                  normalizedImages[0] = { ...normalizedImages[0], isDefault: true };
                }
                const rows = await db
                  .select({ genre: genres.genre })
                  .from(genres)
                  .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
                  .where(eq(genresBooks.bookId, id));

                const databaseGenres = rows.map((row) => row.genre);

                blocks.push({
                  type: 'book',
                  images: normalizedImages.map((image) => ({
                    ...image,
                    selected: image.isDefault,
                    spineColor: image.spineColor,
                  })),
                  blockInfo: {
                    title: foundTitle,
                    author: foundAuthor,
                    pubYear,
                    pageCount,
                    spineColor:
                      normalizedImages.find((image) => image.isDefault)?.spineColor ??
                      normalizedImages[0]?.spineColor ??
                      spineColor,
                    genres: databaseGenres,
                  },
                  blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                  isDatabase: true,
                });
              } else {
                const { title: foundTitle, spineColor } = foundMedia;
                const normalizedImages =
                  'images' in foundMedia && Array.isArray(foundMedia.images)
                    ? foundMedia.images
                    : ('imageUrls' in foundMedia && Array.isArray(foundMedia.imageUrls)
                        ? foundMedia.imageUrls.map((url) => ({
                            url,
                            isDefault: false,
                            spineColor,
                            selected: false,
                          }))
                        : []);
                if (normalizedImages.length > 0 && !normalizedImages.some((image) => image.isDefault)) {
                  normalizedImages[0] = { ...normalizedImages[0], isDefault: true };
                }

                blocks.push({
                  type,
                  images: normalizedImages.map((image) => ({
                    ...image,
                    selected: image.isDefault,
                    spineColor: image.spineColor,
                  })),
                  blockInfo: {
                    title: foundTitle,
                    spineColor:
                      normalizedImages.find((image) => image.isDefault)?.spineColor ??
                      normalizedImages[0]?.spineColor ??
                      spineColor,
                    genres: [],
                  },
                  blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                  isDatabase: true,
                });
              }
            }
          } else {
            //if the media wasnt in the database, collect cover images
            const images = await getMediaCovers(title, author, type);
            // Count each Google API lookup; persist atomically after collection.
            googleApiQueryIncrement += 1;
            if (type === 'book') {
              const bookInfo = await getOpenLibraryData(title, author);
              blocks.push({
                type,
                images: images.map((url, index) => ({
                  url,
                  selected: index === 0,
                  isDefault: index === 0,
                  spineColor: '#ffffff',
                })),
                blockInfo: { ...bookInfo, spineColor: '#ffffff', genres: [] },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: false,
              });
            } else {
              blocks.push({
                type,
                images: images.map((url, index) => ({
                  url,
                  selected: index === 0,
                  isDefault: index === 0,
                  spineColor: '#ffffff',
                })),
                blockInfo: { title, spineColor: '#ffffff', genres: [] },
                blockID: `BLK-${Math.random().toString(36).slice(2, 10).toUpperCase()}`,
                isDatabase: false,
              });
            }
          }
        }
      }
      if (googleApiQueryIncrement > 0) {
        await db
          .insert(googleApiQueryUsage)
          .values({
            date: todayStr,
            queryCount: googleApiQueryIncrement,
          })
          .onConflictDoUpdate({
            target: googleApiQueryUsage.date,
            set: {
              queryCount: sql`${googleApiQueryUsage.queryCount} + ${googleApiQueryIncrement}`,
            },
          });
      }
      const databaseSorted = [...blocks].sort((a, b) =>
        a.isDatabase === b.isDatabase ? 0 : a.isDatabase ? 1 : -1,
      );
      return databaseSorted;
    }),
});
