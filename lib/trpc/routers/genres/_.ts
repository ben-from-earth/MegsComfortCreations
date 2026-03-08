import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import Genre from 'lib/database/models/genre';
import { Genre as GenreEnum } from '@/lib/enums/genreEnums';
import type {
  BookRow,
  SuccessfulGenreLinkUnlinkResponse,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
// database import
import { db } from '@/db/client';

// interfaces and types
import { genres, genresBooks } from '@/db/schema';
import { eq } from 'drizzle-orm';

export const genresRouter = router({
  getAll: adminProcedure.query(async () => {
    const genres = await Genre.getAllGenres();
    return { message: 'Success', genres } as {
      message: string;
      genres: GenreEnum[];
    };
  }),

  getForBook: adminProcedure
    .input(z.object({ bookID: z.uuid() }))
    .query(async ({ input }) => {
      const genres = await Genre.getforbook(input.bookID);
      return {
        message: `Successfully grabbed genres for bookID ${input.bookID}`,
        genres,
      };
    }),

  link: adminProcedure
    .input(
      z.object({
        bookID: z.uuid(),
        genres: z.array(z.string().min(1)),
      }),
    )
    .mutation(async ({ input }) => {
      const responses: SuccessfulGenreLinkUnlinkResponse[] = [];
      for (const genre of input.genres) {
        const [genreRow] = await db
          .select({ id: genres.id })
          .from(genres)
          .where(eq(genres.genre, genre));

        if (!genreRow) {
          responses.push({
            message: `Genre "${genre}" not found in database.`,
            genre,
            bookID: input.bookID,
          });
          continue;
        }

        // 2. Insert into join table
        await db.insert(genresBooks).values({
          bookId: input.bookID,
          genreId: genreRow.id,
        });
        responses.push({
          message: 'Successful genre link',
          genre,
          bookID: input.bookID,
        });
      }
      return { genreResponses: responses };
    }),

  unlink: adminProcedure
    .input(
      z.object({
        bookID: z.uuid(),
        genres: z.array(z.string().min(1)),
      }),
    )
    .mutation(async ({ input }) => {
      const responses: SuccessfulGenreLinkUnlinkResponse[] = [];
      for (const g of input.genres) {
        await Genre.unlink(g, input.bookID);
        responses.push({
          message: 'Successful genre unlink',
          genre: g,
          bookID: input.bookID,
        });
      }
      return { genreResponses: responses };
    }),

  paginateByGenre: adminProcedure
    .input(
      z.object({
        genre: z.string().min(1),
        limit: z.number().int().positive(),
        page: z.number().int().positive(),
        sort: z.enum(['title', 'pubYear', 'spineColor']),
        ascDesc: z.enum(['asc', 'desc']),
      }),
    )
    .query(async ({ input }) => {
      const offset = (input.page - 1) * input.limit;
      const genreRes: { books: BookRow[]; total: number } =
        await Genre.getBooksWithGenre(
          input.genre,
          input.sort,
          offset,
          input.limit,
          input.ascDesc,
        );
      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList: genreRes.books,
        total: genreRes.total,
      };
      return res;
    }),

  paginateNoGenres: adminProcedure
    .input(
      z.object({
        limit: z.number().int().positive(),
        page: z.number().int().positive(),
        sort: z.enum(['title', 'pubYear', 'spineColor']),
        ascDesc: z.enum(['asc', 'desc']),
      }),
    )
    .query(async ({ input }) => {
      const offset = (input.page - 1) * input.limit;
      const genreRes = await Genre.getNoGenreBooks(
        input.sort,
        offset,
        input.limit,
        input.ascDesc,
      );
      const res: SuccessfulPaginationResponse = {
        message: 'Successful database gather',
        paginatedList: genreRes.books,
        total: genreRes.total,
      };
      return res;
    }),
});
