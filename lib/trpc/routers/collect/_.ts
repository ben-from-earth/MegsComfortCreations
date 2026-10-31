import { adminProcedure, router } from 'lib/trpc/trpc';
import { TRPCError } from '@trpc/server';
import { eq, sql } from 'drizzle-orm';
import { z } from 'zod';
import { db as defaultDb, type Db } from '@/db/client';
import { genres, genresBooks } from '@/db/schema';
import { MEDIA_TYPES, type MediaType } from 'lib/constants/media-types';

import { getOpenLibraryData } from './actions/get-open-library-data';
import { getMediaCovers } from './actions/get-media-covers';
import { searchByTitle } from '../database/actions/search-by-title';
import { googleApiQueryUsage } from '@/db/schema';
import {
  convertMediaItemToForm,
  type MediaItemForm,
} from '@/mediacollector/collector-form/media-item-form-schema';
import type { CollectorSearchProgress } from '@/mediacollector/collector-search-progress';
import { persistUploadedImageToS3 } from 'lib/media-storage/local-image-storage';
import { fetchImageBufferAxios } from 'lib/helpers/output-png';
import { extractSpinePalette } from 'lib/helpers/spine-palette';
import sharp from 'sharp';

const SPINE_PALETTE_SAMPLE_SIZE = 64;
const SPINE_PALETTE_CONCURRENCY = 8;
const SPINE_PALETTE_FETCH_TIMEOUT_MS = 8_000;
const UNSET_SPINE_COLOR = '#ffffff';

async function readSpinePalette(image: Buffer | string): Promise<string[]> {
  try {
    const imageBuffer =
      typeof image === 'string'
        ? await fetchImageBufferAxios(image, {
            retries: 0,
            timeout: SPINE_PALETTE_FETCH_TIMEOUT_MS,
          })
        : image;
    const { data, info } = await sharp(imageBuffer)
      .toColourspace('srgb')
      .resize(SPINE_PALETTE_SAMPLE_SIZE, SPINE_PALETTE_SAMPLE_SIZE, {
        fit: 'inside',
      })
      .ensureAlpha()
      .raw()
      .toBuffer({ resolveWithObject: true });
    return extractSpinePalette(data, info.channels);
  } catch (error) {
    console.warn('[readSpinePalette] unable to read cover colors', error);
    return [];
  }
}

async function addSpinePalettes(blocks: MediaItemForm[]) {
  const pendingUrls = [
    ...new Set(
      blocks.flatMap((block) =>
        block.images
          .map((image) => image.url)
          .filter((url) => /^https?:\/\//i.test(url)),
      ),
    ),
  ];
  const palettesByUrl = new Map<string, string[]>();
  await Promise.all(
    Array.from(
      { length: Math.min(SPINE_PALETTE_CONCURRENCY, pendingUrls.length) },
      async () => {
        for (let url = pendingUrls.shift(); url; url = pendingUrls.shift()) {
          palettesByUrl.set(url, await readSpinePalette(url));
        }
      },
    ),
  );

  for (const block of blocks) {
    block.images = block.images.map((image) => {
      const palette = palettesByUrl.get(image.url) ?? [];
      const spineColor =
        !block.isDatabase && image.spineColor === UNSET_SPINE_COLOR
          ? (palette[0] ?? image.spineColor)
          : image.spineColor;
      return { ...image, palette, spineColor };
    });
    if (!block.isDatabase) {
      block.blockInfo.spineColor =
        block.images.find((image) => image.selected)?.spineColor ??
        block.blockInfo.spineColor;
    }
  }
}

async function collectSearchItem(
  db: Db,
  type: MediaType,
  title: string,
  author?: string,
): Promise<{ blocks: MediaItemForm[]; googleApiQueryIncrement: number }> {
  const blocks: MediaItemForm[] = [];
  let googleApiQueryIncrement = 0;

  const mediaSearchData = await searchByTitle(db, type, title, author);

  if (mediaSearchData.total > 0) {
    for (const foundMedia of mediaSearchData.foundMediaList) {
      if (type === 'book' && !('author' in foundMedia)) {
        continue;
      }

      const databaseGenres =
        type === 'book'
          ? (
              await db
                .select({ genre: genres.genre })
                .from(genres)
                .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
                .where(eq(genresBooks.bookId, foundMedia.id))
            ).map((row) => row.genre)
          : [];

      blocks.push(
        convertMediaItemToForm({
          item: {
            id: foundMedia.id,
            title: foundMedia.title,
            spineColor: foundMedia.spineColor,
            images: foundMedia.images,
            ...('author' in foundMedia
              ? {
                  author: foundMedia.author,
                  pageCount: foundMedia.pageCount,
                  pubYear: foundMedia.pubYear,
                }
              : {}),
          },
          type,
          genres: databaseGenres,
        }),
      );
    }
  } else {
    const images = await getMediaCovers(title, author, type);
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
        blockID: crypto.randomUUID(),
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
        blockID: crypto.randomUUID(),
        isDatabase: false,
      });
    }
  }

  return { blocks, googleApiQueryIncrement };
}

async function incrementGoogleApiQueryUsage(
  db: Db,
  googleApiQueryIncrement: number,
  todayStr: string,
) {
  if (googleApiQueryIncrement <= 0) {
    return;
  }

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

function sortDatabaseBlocksLast(blocks: MediaItemForm[]) {
  return [...blocks].sort((a, b) =>
    a.isDatabase === b.isDatabase ? 0 : a.isDatabase ? 1 : -1,
  );
}

const collectMediaInputSchema = z.object({
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
});

type CollectMediaInput = z.infer<typeof collectMediaInputSchema>;

type CollectMediaStreamEvent =
  | { type: 'progress'; progress: CollectorSearchProgress }
  | {
      type: 'result';
      blocks: MediaItemForm[];
      pngExportBlockIds: string[];
    };

function listCollectSearches(input: CollectMediaInput) {
  return MEDIA_TYPES.flatMap((type) =>
    input[type].map((item) => ({
      type,
      title: item.title,
      author: item.author,
    })),
  );
}

async function* collectMediaEvents(
  db: Db,
  input: CollectMediaInput,
): AsyncGenerator<CollectMediaStreamEvent> {
  const searches = listCollectSearches(input);
  const totalCount = searches.length;
  const blocks: MediaItemForm[] = [];
  const todayStr = new Date().toLocaleDateString('en-CA', {
    timeZone: 'America/New_York',
  });
  let googleApiQueryIncrement = 0;

  yield {
    type: 'progress',
    progress: { phase: 'collecting', searchedCount: 0, totalCount },
  };

  for (const [index, search] of searches.entries()) {
    const collected = await collectSearchItem(
      db,
      search.type,
      search.title,
      search.author,
    );
    blocks.push(...collected.blocks);
    googleApiQueryIncrement += collected.googleApiQueryIncrement;
    yield {
      type: 'progress',
      progress: {
        phase: 'collecting',
        searchedCount: index + 1,
        totalCount,
      },
    };
  }

  await incrementGoogleApiQueryUsage(db, googleApiQueryIncrement, todayStr);
  yield {
    type: 'progress',
    progress: { phase: 'pickingColors', isComplete: false },
  };
  await addSpinePalettes(blocks);
  yield {
    type: 'progress',
    progress: { phase: 'pickingColors', isComplete: true },
  };

  const pngExportBlockIds = blocks.map((block) => block.blockID);
  yield {
    type: 'result',
    blocks: sortDatabaseBlocksLast(blocks),
    pngExportBlockIds,
  };
}

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

      const [uploadedImage, palette] = await Promise.all([
        persistUploadedImageToS3({
          imageBuffer: fileBuffer,
          mediaType: 'book',
          mediaId: input.blockID,
          sortOrder: input.sortOrder,
          fileName: input.fileName,
          mimeType: input.mimeType,
        }),
        readSpinePalette(fileBuffer),
      ]);

      return {
        url: uploadedImage.publicPath,
        isDefault: input.sortOrder === 0,
        spineColor: UNSET_SPINE_COLOR,
        selected: true,
        palette,
      };
    }),
  collectMedia: adminProcedure
    .input(collectMediaInputSchema)
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      let result:
        | Extract<CollectMediaStreamEvent, { type: 'result' }>
        | undefined;
      for await (const event of collectMediaEvents(db, input)) {
        if (event.type === 'result') {
          result = event;
        }
      }
      if (!result) {
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: 'Collection finished without a result.',
        });
      }
      return {
        blocks: result.blocks,
        pngExportBlockIds: result.pngExportBlockIds,
      };
    }),
  collectMediaStream: adminProcedure
    .input(collectMediaInputSchema)
    .mutation(async function* ({ input, ctx }) {
      yield* collectMediaEvents(ctx.db ?? defaultDb, input);
    }),
});
