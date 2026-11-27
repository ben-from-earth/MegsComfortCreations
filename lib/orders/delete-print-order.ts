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
): Promise<void> {
  const [order] = await db
    .select({ fileUrl: orders.fileUrl })
    .from(orders)
    .where(eq(orders.id, orderId));
  if (!order) {
    throw new OrderWriteError('not_found', MISSING_ORDER_MESSAGE);
  }

  const key = orderFileObjectKey(order.fileUrl, publicObjectBaseUrl());
  if (key) {
    await deletePublicObject(key);
  }

  const [removed] = await db
    .delete(orders)
    .where(eq(orders.id, orderId))
    .returning({ id: orders.id });
  if (!removed) {
    throw new OrderWriteError('not_found', MISSING_ORDER_MESSAGE);
  }
}
