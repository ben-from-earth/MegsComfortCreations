import { unstable_cache } from 'next/cache';
import { eq } from 'drizzle-orm';
import { db } from '@/db/client';
import { customers, orders } from '@/db/schema';
import { aggregateShippedPlaces } from '@/components/home/shipped-places';

async function readShippedPlaces() {
  const rows = await db
    .select({
      country: customers.country,
      state: customers.state,
    })
    .from(orders)
    .innerJoin(customers, eq(orders.customerId, customers.id));

  return aggregateShippedPlaces(rows);
}

export const getShippedPlaces = unstable_cache(
  readShippedPlaces,
  ['shipped-places'],
  { revalidate: 60 * 60 },
);
