import z from 'zod';
import { mediaItemFormSchema } from './mediaItemFormSchema';
export { mediaItemFormSchema } from './mediaItemFormSchema';
export type { MediaItemForm } from './mediaItemFormSchema';

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
  collectedData: z.array(mediaItemFormSchema),
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
