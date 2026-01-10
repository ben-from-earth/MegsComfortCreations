import z from 'zod';

export const baseBlockInfoSchema = z.object({
  title: z.string(),
  spineColor: z.string(),
  databaseGenres: z.array(z.string()).optional(),
});

export const bookBlockInfoSchema = baseBlockInfoSchema.extend({
  author: z.string().nullable(),
  pubYear: z.number().nullable(),
  pageCount: z.number().nullable(),
});

// export const collectedBlockInformationSchema = z.discriminatedUnion('type', [
//   z.object({
//     type: z.literal('book'),
//     images: z.array(z.string()),
//     blockInfo: bookBlockInfoSchema,
//     blockID: z.string(),
//     isDatabase: z.boolean(),
//   }),
//   z.object({
//     type: z.union([
//       z.literal('movie'),
//       z.literal('videoGame'),
//       z.literal('album'),
//     ]),
//     images: z.array(z.string()),
//     blockInfo: baseBlockInfoSchema,
//     blockID: z.string(),
//     isDatabase: z.boolean(),
//   }),
// ]);

export const collectedBlockInformationSchema = z.object({
  type: z.enum(['book', 'movie', 'videoGame', 'album']),
  images: z.array(z.object({ url: z.string(), selected: z.boolean() })).min(1),
  blockInfo: bookBlockInfoSchema,
  blockID: z.string(),
  isDatabase: z.boolean(),
});

export const collectorFormSchema = z.object({
  orderNumber: z.string(),
  customerName: z.string(),
  collectionList: z.object({
    book: z.array(
      z.object({ title: z.string(), author: z.string().optional() }),
    ),
    // movie: z.array(
    //   z.object({ title: z.string(), author: z.string().optional() }),
    // ),
    // videoGame: z.array(
    //   z.object({ title: z.string(), author: z.string().optional() }),
    // ),
    // album: z.array(
    //   z.object({ title: z.string(), author: z.string().optional() }),
    // ),
  }),
  collectedData: z.array(collectedBlockInformationSchema),
  pngFormat: z.enum(['3', '5']).optional(),
});

export type CollectorFormData = z.infer<typeof collectorFormSchema>;
export type CollectedBlockInformation = z.infer<
  typeof collectorFormSchema
>['collectedData'][number];
