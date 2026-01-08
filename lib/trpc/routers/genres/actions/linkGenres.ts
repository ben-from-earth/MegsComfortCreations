import { SuccessfulGenreLinkUnlinkResponse } from 'lib/interfaces/globalInterfaces';
import Genre from 'lib/database/models/genre';

export async function linkGenres(bookID: string, genres: string[]) {
  const responses: SuccessfulGenreLinkUnlinkResponse[] = [];
  for (const g of genres) {
    await Genre.link(g, bookID);
    responses.push({
      message: 'Successful genre link',
      genre: g,
      bookID: bookID,
    });
  }
  return { genreResponses: responses };
}
