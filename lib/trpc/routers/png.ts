import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import { outputAuto } from 'lib/helpers/outputPNG';

const templateSchema = z.union([z.literal(3), z.literal(5)]);
const imageSchema = z.object({
  url: z.string().url(),
  type: z.enum(['book', 'movie', 'videoGame', 'album']),
  spineColor: z.string().min(1),
});

export const pngRouter = router({
  create: adminProcedure
    .input(
      z.object({
        template: templateSchema,
        images: z.array(imageSchema),
        customerName: z.string(),
        orderNumber: z.string(),
        repeatCount: z.number().min(1),
      }),
    )
    .mutation(async ({ input }) => {
      const firstName = input.customerName.split(' ')[0] || 'Customer';
      const lastInititial = input.customerName.split(' ')[1]
        ? input.customerName.split(' ')[1][0]
        : 'NoLastInitial';
      input.images = Array.from(
        { length: input.repeatCount },
        () => input.images,
      ).flat();
      const { mime, filename, buffer } = await outputAuto({
        template: input.template,
        images: input.images,
        fileOutputName: `${firstName}_${lastInititial}_${input.orderNumber}`,
      });
      const dataBase64 = Buffer.from(buffer).toString('base64');
      return { mime, filename, dataBase64 };
    }),
});
