import z from 'zod';

export const baseBlockInfoSchema = z.object({
  title: z.string(),
  spineColor: z.string(),
  genres: z.array(z.string()),
});

export const bookBlockInfoSchema = baseBlockInfoSchema.extend({
  author: z.string().nullable(),
  pubYear: z.number().nullable(),
  pageCount: z.number().nullable(),
});

export const otherMediaBlockInfoSchema = baseBlockInfoSchema;

const imageSelectionSchema = z.object({
  url: z.string(),
  selected: z.boolean(),
});

export const collectedBlockInformationSchema = z.object({
  type: z.enum(['book', 'movie', 'videoGame', 'album']),
  images: z.array(imageSelectionSchema).min(1),
  blockInfo: baseBlockInfoSchema.extend({
    author: z.string().nullable().optional(),
    pubYear: z.number().nullable().optional(),
    pageCount: z.number().nullable().optional(),
  }),
  blockID: z.string(),
  isDatabase: z.boolean(),
});

export const collectorFormSchema = z.object({
  orderNumber: z.string(),
  customerName: z.string(),
  bookClubRepeat: z.number(),
  collectionList: z.object({
    book: z.array(
      z.object({ title: z.string(), author: z.string().optional() }),
    ),
    movie: z.array(
      z.object({ title: z.string(), author: z.string().optional() }),
    ),
    videoGame: z.array(
      z.object({ title: z.string(), author: z.string().optional() }),
    ),
    album: z.array(
      z.object({ title: z.string(), author: z.string().optional() }),
    ),
  }),
  collectedData: z.array(collectedBlockInformationSchema),
  pngFormat: z.enum(['3', '5']).optional(),
});

export type CollectorFormData = z.infer<typeof collectorFormSchema>;
export type CollectedBlockInformation = z.infer<
  typeof collectorFormSchema
>['collectedData'][number];
