import { router, adminProcedure } from 'lib/trpc/trpc';
import { z } from 'zod';
import { Genre as GenreEnum } from '@/lib/enums/genre-enums';
import type { SuccessfulGenreLinkUnlinkResponse } from 'lib/interfaces/global-interfaces';
import { db as defaultDb } from '@/db/client';
import { genres, genresBooks } from '@/db/schema';
import { and, eq } from 'drizzle-orm';

export const genresRouter = router({
  getAll: adminProcedure.query(async ({ ctx }) => {
    const db = ctx.db ?? defaultDb;
    const rows = await db.select({ genre: genres.genre }).from(genres);
    const genreList = rows.map((row) => row.genre);
    return { message: 'Success', genres: genreList } as {
      message: string;
      genres: GenreEnum[];
    };
  }),

  getForBook: adminProcedure
    .input(z.object({ bookID: z.uuid() }))
    .query(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const rows = await db
        .select({ genre: genres.genre })
        .from(genres)
        .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
        .where(eq(genresBooks.bookId, input.bookID));
      const genreList = rows.map((row) => row.genre);
      return {
        message: `Successfully grabbed genres for bookID ${input.bookID}`,
        genres: genreList,
      };
    }),

  link: adminProcedure
    .input(
      z.object({
        bookID: z.uuid(),
        genres: z.array(z.string().min(1)),
      }),
    )
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
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
    .mutation(async ({ input, ctx }) => {
      const db = ctx.db ?? defaultDb;
      const responses: SuccessfulGenreLinkUnlinkResponse[] = [];
      for (const g of input.genres) {
        const [genreRow] = await db
          .select({ id: genres.id })
          .from(genres)
          .where(eq(genres.genre, g));

        if (genreRow) {
          await db
            .delete(genresBooks)
            .where(
              and(
                eq(genresBooks.bookId, input.bookID),
                eq(genresBooks.genreId, genreRow.id),
              ),
            );
        }
        responses.push({
          message: 'Successful genre unlink',
          genre: g,
          bookID: input.bookID,
        });
      }
      return { genreResponses: responses };
    }),
});
