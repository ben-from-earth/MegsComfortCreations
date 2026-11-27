import type { DbExecutor } from '@/db/client';
import { orders } from '@/db/schema';
import { eq } from 'drizzle-orm';
import {
  deletePublicObject,
  publicObjectBaseUrl,
} from 'lib/media-storage/local-image-storage';
import { OrderWriteError } from './save-print-order';
import { orderFileObjectKey } from './order-file-object-key';

const MISSING_ORDER_MESSAGE = 'This order no longer exists.';

export async function deletePrintOrder(
  db: DbExecutor,
  orderId: string,
): Promise<{ fileDeleted: boolean }> {
  const [order] = await db
    .select({ id: orders.id, fileUrl: orders.fileUrl })
    .from(orders)
    .where(eq(orders.id, orderId))
    .limit(1);
  if (!order) {
    throw new OrderWriteError('not_found', MISSING_ORDER_MESSAGE);
  }

  const deleted = await db
    .delete(orders)
    .where(eq(orders.id, orderId))
    .returning({ id: orders.id });
  if (deleted.length === 0) {
    throw new OrderWriteError('not_found', MISSING_ORDER_MESSAGE);
  }

  try {
    const key = orderFileObjectKey(order.fileUrl, publicObjectBaseUrl());
    if (!key) {
      return { fileDeleted: true };
    }
    await deletePublicObject(key);
    return { fileDeleted: true };
  } catch (error) {
    console.error(
      '[deletePrintOrder] order removed, print file remains',
      error instanceof Error ? error.message : String(error),
    );
    return { fileDeleted: false };
  }
}
