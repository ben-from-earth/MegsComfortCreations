import z from 'zod';

export const collectorFormSchema = z.object({
  orderNumber: z.string(),
  customerName: z.string(),
  collectionData: z.object({
    books: z.array(
      z.object({
        title: z.string(),
        author: z.string(),
        pageCount: z.number().optional(),
        publicationYear: z.number().optional(),
        images: z.array(
          z.object({
            src: z.string(),
            selected: z.boolean(),
          }),
        ),
        genres: z.array(z.string()),
      }),
    ),
    movies: z.array(
      z.object({
        title: z.string(),
        images: z.array(
          z.object({
            src: z.string(),
            selected: z.boolean(),
          }),
        ),
      }),
    ),
    videoGames: z.array(
      z.object({
        title: z.string(),
        images: z.array(
          z.object({
            src: z.string(),
            selected: z.boolean(),
          }),
        ),
      }),
    ),
    albums: z.array(
      z.object({
        title: z.string(),
        images: z.array(
          z.object({
            src: z.string(),
            selected: z.boolean(),
          }),
        ),
      }),
    ),
  }),
});

export type CollectorFormData = z.infer<typeof collectorFormSchema>;
