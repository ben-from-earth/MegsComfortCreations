import { TRPCError } from '@trpc/server';
import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import { outputAuto } from 'lib/helpers/output-png';
import {
  OrderWriteError,
  resolvePrintOrder,
  storePrintOrder,
  type ResolvedPrintOrder,
} from 'lib/orders/save-print-order';

const templateSchema = z.union([z.literal(3), z.literal(5)]);
const orderPayloadSchema = z.object({
  customerName: z.string().trim().min(1, 'Customer name is required.'),
  orderNumber: z.string().trim().min(1, 'Order number is required.'),
  template: templateSchema,
  repeatCount: z.number().int().min(1).max(25),
  coverIds: z.array(z.uuid()),
  orderId: z.uuid().optional(),
});
const imageSchema = z.object({
  url: z.string().url(),
  type: z.enum(['book', 'movie', 'videoGame', 'album']),
  spineColor: z.string().min(1),
  codeUrl: z.string().url().optional(),
  console: z.string().optional(),
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
        order: orderPayloadSchema.optional(),
      }),
    )
    .mutation(async ({ input, ctx }) => {
      let resolvedOrder: ResolvedPrintOrder | null = null;
      if (input.order) {
        try {
          resolvedOrder = await resolvePrintOrder(ctx.db, input.order);
        } catch (error) {
          rethrowOrderWrite(error);
        }
      }

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
      if (resolvedOrder) {
        try {
          await storePrintOrder(ctx.db, resolvedOrder, {
            mimeType: mime,
            fileName: filename,
            body: Buffer.from(buffer),
          });
        } catch (error) {
          rethrowOrderWrite(error);
        }
      }

      const dataBase64 = Buffer.from(buffer).toString('base64');
      return { mime, filename, dataBase64 };
    }),
});

function rethrowOrderWrite(error: unknown): never {
  if (error instanceof OrderWriteError) {
    throw new TRPCError({
      code:
        error.reason === 'conflict'
          ? 'CONFLICT'
          : error.reason === 'not_found'
            ? 'NOT_FOUND'
            : 'BAD_REQUEST',
      message: error.message,
    });
  }
  throw error;
}
