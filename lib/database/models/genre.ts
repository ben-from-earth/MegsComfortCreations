// database import
import { db } from '@/db/client';

// interfaces and types
import { BookRow } from 'lib/interfaces/globalInterfaces';
import { books, genres, genresBooks } from '@/db/schema';
import { and, asc, desc, eq, isNull, sql } from 'drizzle-orm';

const validSortKeys = ['title', 'pubYear', 'spineColor'] as const;
type SortKey = (typeof validSortKeys)[number];

export default class Genre {
  static async getAllGenres(): Promise<string[]> {
    const rows = await db.select({ genre: genres.genre }).from(genres);

    return rows.map((row) => row.genre);
  }

  static async link(genre: string, bookID: string): Promise<void> {
    // 1. Find genre id
    const [genreRow] = await db
      .select({ id: genres.id })
      .from(genres)
      .where(eq(genres.genre, genre));

    if (!genreRow) {
      // previous raw SQL would just crash; you can throw a nicer error if you like
      throw new Error(`Genre "${genre}" not found`);
    }

    // 2. Insert into join table
    await db.insert(genresBooks).values({
      bookId: bookID,
      genreId: genreRow.id,
    });
  }

  static async unlink(genre: string, bookID: string): Promise<void> {
    // simplest: find genreId first, then delete that link
    const [genreRow] = await db
      .select({ id: genres.id })
      .from(genres)
      .where(eq(genres.genre, genre));

    if (!genreRow) {
      // nothing to unlink
      return;
    }

    await db
      .delete(genresBooks)
      .where(
        and(
          eq(genresBooks.bookId, bookID),
          eq(genresBooks.genreId, genreRow.id),
        ),
      );
  }

  static async getforbook(bookID: string): Promise<string[]> {
    const rows = await db
      .select({ genre: genres.genre })
      .from(genres)
      .innerJoin(genresBooks, eq(genresBooks.genreId, genres.id))
      .where(eq(genresBooks.bookId, bookID));

    return rows.map((row) => row.genre);
  }

  static async getNoGenreBooks(
    sort: string,
    offset: number,
    limit: number,
    ascDesc: string,
  ): Promise<{ books: BookRow[]; total: number }> {
    const sortKey = sort as SortKey;
    const direction = ascDesc.toLowerCase() === 'desc' ? 'desc' : 'asc';

    if (!validSortKeys.includes(sortKey)) {
      throw new Error(
        `Invalid sort key: ${sort}. Must be one of ${validSortKeys.join(', ')}`,
      );
    }

    // map sort key string to actual column
    const sortColumn =
      sortKey === 'title'
        ? books.title
        : sortKey === 'pubYear'
          ? books.pubYear
          : books.spineColor;

    const orderByExpr =
      direction === 'desc' ? desc(sortColumn) : asc(sortColumn);

    const rows = await db
      .select({ book: books })
      .from(books)
      .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
      .where(isNull(genresBooks.bookId))
      .orderBy(orderByExpr)
      .limit(limit)
      .offset(offset);

    const [{ value: total }] = await db
      .select({ value: sql<number>`count(*)` })
      .from(books)
      .leftJoin(genresBooks, eq(genresBooks.bookId, books.id))
      .where(isNull(genresBooks.bookId));

    return {
      books: rows.map((r) => r.book) as BookRow[],
      total,
    };
  }

  static async getBooksWithGenre(
    genre: string,
    sort: string,
    offset: number,
    limit: number,
    ascDesc: string,
  ): Promise<{ books: BookRow[]; total: number }> {
    const sortKey = sort as SortKey;
    const direction = ascDesc.toLowerCase() === 'desc' ? 'desc' : 'asc';

    if (!validSortKeys.includes(sortKey)) {
      throw new Error(
        `Invalid sort key: ${sort}. Must be one of ${validSortKeys.join(', ')}`,
      );
    }

    const sortColumn =
      sortKey === 'title'
        ? books.title
        : sortKey === 'pubYear'
          ? books.pubYear
          : books.spineColor;

    const orderByExpr =
      direction === 'desc' ? desc(sortColumn) : asc(sortColumn);

    const rows = await db
      .select({ book: books })
      .from(books)
      .innerJoin(genresBooks, eq(genresBooks.bookId, books.id))
      .innerJoin(genres, eq(genres.id, genresBooks.genreId))
      .where(eq(genres.genre, genre))
      .orderBy(orderByExpr)
      .limit(limit)
      .offset(offset);

    const [{ value: total }] = await db
      .select({ value: sql<number>`count(*)` })
      .from(books)
      .innerJoin(genresBooks, eq(genresBooks.bookId, books.id))
      .innerJoin(genres, eq(genres.id, genresBooks.genreId))
      .where(eq(genres.genre, genre));

    return {
      books: rows.map((r) => r.book) as BookRow[],
      total,
    };
  }
}
