import { randomUUID } from 'node:crypto';
import type { DbExecutor } from '@/db/client';
import { mediaImageItems, orderCovers, orders } from '@/db/schema';
import { eq, inArray } from 'drizzle-orm';
import { putOrderFile } from 'lib/media-storage/order-file-storage';
import { chooseOrderWrite } from './choose-order-write';

const MISSING_ORDER_MESSAGE = 'This order no longer exists.';
const ORDER_NUMBER_CONFLICT_MESSAGE =
  'That order number is already used by another order.';
const MISSING_COVER_MESSAGE = 'A cover on this order no longer exists.';

export class OrderWriteError extends Error {
  readonly reason: 'bad_request' | 'conflict' | 'not_found';

  constructor(
    reason: 'bad_request' | 'conflict' | 'not_found',
    message: string,
  ) {
    super(message);
    this.name = 'OrderWriteError';
    this.reason = reason;
  }
}

export type PrintOrderInput = {
  customerName: string;
  orderNumber: string;
  template: 3 | 5;
  repeatCount: number;
  coverIds: string[];
  orderId?: string;
};

export type ResolvedPrintOrder = {
  orderId: string;
  action: 'insert' | 'update';
  customerName: string;
  orderNumber: string;
  template: 3 | 5;
  repeatCount: number;
  coverIds: string[];
};

function errorText(error: unknown) {
  const parts: string[] = [];
  const visit = (value: unknown, depth: number) => {
    if (depth > 4 || value === null || value === undefined) {
      return;
    }
    if (value instanceof Error) {
      parts.push(value.message);
      visit(value.cause, depth + 1);
    }
    if (typeof value === 'object' && value !== null) {
      if ('constraint' in value && typeof value.constraint === 'string') {
        parts.push(value.constraint);
      }
      if ('code' in value && typeof value.code === 'string') {
        parts.push(value.code);
      }
    }
  };
  visit(error, 0);
  return parts.join(' ');
}

function unwrapOrderWriteError(error: unknown) {
  let current: unknown = error;
  for (let depth = 0; depth < 5 && current instanceof Error; depth += 1) {
    if (current instanceof OrderWriteError) {
      return current;
    }
    current = current.cause;
  }
  return null;
}

async function loadOrderById(db: DbExecutor, id: string) {
  const [row] = await db
    .select({ id: orders.id, orderNumber: orders.orderNumber })
    .from(orders)
    .where(eq(orders.id, id))
    .limit(1);
  return row ?? null;
}

async function loadOrderByNumber(db: DbExecutor, orderNumber: string) {
  const [row] = await db
    .select({ id: orders.id, orderNumber: orders.orderNumber })
    .from(orders)
    .where(eq(orders.orderNumber, orderNumber))
    .limit(1);
  return row ?? null;
}

async function hasMissingCover(db: DbExecutor, coverIds: readonly string[]) {
  const uniqueIds = [...new Set(coverIds)];
  if (uniqueIds.length === 0) {
    return false;
  }
  const found = await db
    .select({ id: mediaImageItems.id })
    .from(mediaImageItems)
    .where(inArray(mediaImageItems.id, uniqueIds));
  return found.length !== uniqueIds.length;
}

export async function resolvePrintOrder(
  db: DbExecutor,
  input: PrintOrderInput,
): Promise<ResolvedPrintOrder> {
  const customerName = input.customerName.trim();
  const orderNumber = input.orderNumber.trim();
  const requestedOrderId = input.orderId ?? null;
  const [orderWithRequestedId, orderWithRequestedNumber] = await Promise.all([
    requestedOrderId
      ? loadOrderById(db, requestedOrderId)
      : Promise.resolve(null),
    loadOrderByNumber(db, orderNumber),
  ]);
  const choice = chooseOrderWrite({
    requestedOrderId,
    orderWithRequestedId,
    orderWithRequestedNumber,
  });
  if (choice.action === 'missing') {
    throw new OrderWriteError('not_found', MISSING_ORDER_MESSAGE);
  }
  if (choice.action === 'conflict') {
    throw new OrderWriteError('conflict', ORDER_NUMBER_CONFLICT_MESSAGE);
  }
  if (await hasMissingCover(db, input.coverIds)) {
    throw new OrderWriteError('bad_request', MISSING_COVER_MESSAGE);
  }
  return {
    orderId: choice.action === 'update' ? choice.orderId : randomUUID(),
    action: choice.action,
    customerName,
    orderNumber,
    template: input.template,
    repeatCount: input.repeatCount,
    coverIds: input.coverIds,
  };
}

export async function storePrintOrder(
  db: DbExecutor,
  order: ResolvedPrintOrder,
  file: { mimeType: 'image/png' | 'application/zip'; fileName: string; body: Buffer },
): Promise<void> {
  const storedFile = await putOrderFile({
    orderId: order.orderId,
    fileName: file.fileName,
    mimeType: file.mimeType,
    body: file.body,
  });

  try {
    await db.transaction(async (tx) => {
      if (await hasMissingCover(tx, order.coverIds)) {
        throw new OrderWriteError('bad_request', MISSING_COVER_MESSAGE);
      }
      if (order.action === 'insert') {
        await tx.insert(orders).values({
          id: order.orderId,
          customerName: order.customerName,
          orderNumber: order.orderNumber,
          pngFormat: order.template,
          repeatCount: order.repeatCount,
          fileUrl: storedFile.fileUrl,
          fileName: storedFile.fileName,
          mimeType: storedFile.mimeType,
        });
      } else {
        const [updated] = await tx
          .update(orders)
          .set({
            customerName: order.customerName,
            orderNumber: order.orderNumber,
            pngFormat: order.template,
            repeatCount: order.repeatCount,
            fileUrl: storedFile.fileUrl,
            fileName: storedFile.fileName,
            mimeType: storedFile.mimeType,
            updatedAt: new Date(),
          })
          .where(eq(orders.id, order.orderId))
          .returning({ id: orders.id });
        if (!updated) {
          throw new OrderWriteError('not_found', MISSING_ORDER_MESSAGE);
        }
      }

      await tx.delete(orderCovers).where(eq(orderCovers.orderId, order.orderId));
      if (order.coverIds.length > 0) {
        await tx.insert(orderCovers).values(
          order.coverIds.map((mediaImageItemId, position) => ({
            orderId: order.orderId,
            mediaImageItemId,
            position,
          })),
        );
      }
    });
  } catch (error) {
    const orderWriteError = unwrapOrderWriteError(error);
    if (orderWriteError) {
      throw orderWriteError;
    }
    const hint = errorText(error);
    if (hint.includes('orders_order_number_unique')) {
      throw new OrderWriteError('conflict', ORDER_NUMBER_CONFLICT_MESSAGE);
    }
    if (hint.includes('order_covers_media_image_item_id')) {
      throw new OrderWriteError('bad_request', MISSING_COVER_MESSAGE);
    }
    throw error;
  }
}
