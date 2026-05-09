import { and, eq, ilike } from 'drizzle-orm';
import { Db } from '@/db/client';
import { MediaType } from 'lib/constants/mediaTypes';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { books, otherMedia } from '@/db/schema';
import {
  loadBookImagesById,
  loadOtherMediaImagesById,
} from 'lib/media-storage/media-image-records';

export async function searchByTitle(db: Db, type: MediaType, title: string) {
  const rearrangedTitle = titleRearrange(title);
  const result =
    type === 'book'
      ? await db.select().from(books).where(ilike(books.title, rearrangedTitle))
      : await (() => {
          const otherType = type;
          return db
            .select()
            .from(otherMedia)
            .where(
              and(
                eq(otherMedia.mediaType, otherType),
                ilike(otherMedia.title, rearrangedTitle),
              ),
            );
        })();
  const normalizedResult =
    type === 'book'
      ? await (async () => {
          const imagesByBookId = await loadBookImagesById(
            db,
            result.map((row) => row.id),
          );
          return result.map((row) => ({
            ...row,
            images: imagesByBookId.get(row.id) ?? [],
          }));
        })()
      : await (async () => {
          const imagesByOtherMediaId = await loadOtherMediaImagesById(
            db,
            result.map((row) => row.id),
          );
          return result.map((row) => ({
            ...row,
            images: imagesByOtherMediaId.get(row.id) ?? [],
          }));
        })();
  const total = normalizedResult.length;

  const message =
    total === 0
      ? `No ${type} in database with title ${rearrangedTitle}`
      : `Successfully found ${total} ${type}(s) with title ${titleRearrange(normalizedResult[0].title)}`;

  return {
    message,
    foundMediaList: normalizedResult,
    total,
  };
}
