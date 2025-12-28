import { router, publicProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import { outputAuto, type Template } from 'lib/helpers/outputPNG';
import type { ImageData } from 'lib/state/slices/pngCollectionSlice';

const templateSchema = z.union([z.literal(3), z.literal(5)]);
const imageSchema = z.object({
  url: z.string().url(),
  type: z.enum(['book', 'movie', 'videoGame', 'album']),
  spineColor: z.string().min(1),
});

export const pngRouter = router({
  create: publicProcedure
    .input(z.object({ template: templateSchema, images: z.array(imageSchema) }))
    .mutation(async ({ input }) => {
      const { mime, filename, buffer } = await outputAuto({
        template: input.template as Template,
        images: input.images as ImageData[],
        prefix: 'grid',
      });
      const dataBase64 = Buffer.from(buffer).toString('base64');
      return { mime, filename, dataBase64 };
    }),
});
