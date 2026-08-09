import z from 'zod';

export const baseBlockInfoSchema = z.object({
  title: z.string(),
  spineColor: z.string(),
  genres: z.array(z.string()),
});

const imageSelectionSchema = z.object({
  url: z.string(),
  selected: z.boolean(),
  isDefault: z.boolean(),
  spineColor: z.string(),
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
}).superRefine((value, ctx) => {
  const selectedCount = value.images.filter((image) => image.selected).length;
  if (selectedCount > 1) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: 'Only one image can be selected at a time.',
      path: ['images'],
    });
  }
});

export const collectorFormSchema = z.object({
  orderNumber: z.string(),
  customerName: z.string(),
  bookClubRepeat: z
    .number()
    .min(1, 'Book Club Repeat Number must be at least 1.'),
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
  pngFormat: z
    .enum(['3', '5'], {
      error: 'Please select a PNG template option',
    })
    .nullable()
    .refine((value) => value !== null, {
      message: 'Please select a PNG template option',
    }),
});

export type CollectorFormData = z.infer<typeof collectorFormSchema>;
export type CollectedBlockInformation = z.infer<
  typeof collectorFormSchema
>['collectedData'][number];
