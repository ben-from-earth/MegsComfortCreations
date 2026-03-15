import { and, eq, ilike } from 'drizzle-orm';
import { Db } from '@/db/client';
import { MediaType } from 'lib/interfaces/globalInterfaces';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { books, otherMedia } from '@/db/schema';

export async function searchByTitle(db: Db, type: MediaType, title: string) {
  const rearrangedTitle = titleRearrange(title);
  const result =
    type === 'book'
      ? await db
          .select()
          .from(books)
          .where(ilike(books.title, rearrangedTitle))
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
  const total = result.length;

  const message =
    total === 0
      ? `No ${type} in database with title ${rearrangedTitle}`
      : `Successfully found ${total} ${type}(s) with title ${titleRearrange(result[0].title)}`;

  return {
    message,
    foundMediaList: result,
    total,
  };
}
