import { ilike } from 'drizzle-orm';
import { Db } from '@/db/client';
import { MediaType } from 'lib/interfaces/globalInterfaces';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { albums, books, movies, videoGames } from '@/db/schema';

export async function searchByTitle(db: Db, type: MediaType, title: string) {
  const tableMap = {
    book: books,
    movie: movies,
    videoGame: videoGames,
    album: albums,
  } as const;

  const table = tableMap[type];
  const rearrangedTitle = titleRearrange(title);
  const result = await db
    .select()
    .from(table)
    .where(ilike(table.title, rearrangedTitle));
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
