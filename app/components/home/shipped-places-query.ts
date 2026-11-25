import { unstable_cache } from 'next/cache';
import { aggregateShippedPlaces } from '@/components/home/shipped-places';

async function readShippedPlaces() {
  return aggregateShippedPlaces([]);
}

export const getShippedPlaces = unstable_cache(
  readShippedPlaces,
  ['shipped-places'],
  { revalidate: 60 * 60 },
);
