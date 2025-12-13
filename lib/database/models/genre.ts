// database import
import { db } from '@/app/db/client';

// interfaces and types
import { PostSavedMediaItem } from '@/lib/interfaces/globalInterfaces';

export default class Genre {
  static async getAllGenres() {
    const dbRes = await db.query<{ genre: string }>(`SELECT genre FROM genres`);
    const result: string[] = dbRes.rows.map((row) => row.genre);
    return result;
  }

  static async link(genre: string, bookID: string): Promise<void> {
    const genreFind = await db.query<{ id: string; genre: string }>(
      `SELECT id FROM genres
      WHERE genre=$1`,
      [genre],
    );

    const genreID = genreFind.rows[0].id;
    await db.query(
      `INSERT INTO genres_books (
        book_id, genre_id
      ) VALUES ($1, $2)`,
      [bookID, genreID],
    );
  }

  static async unlink(genre: string, bookID: string): Promise<void> {
    await db.query(
      `DELETE FROM genres_books gb
        USING genres g
        WHERE gb.book_id = $1
        AND g.genre = $2
        AND gb.genre_id = g.id`,
      [bookID, genre],
    );
  }

  static async getforbook(bookID: string): Promise<string[]> {
    const result = await db.query<{ genre: string }>(
      `SELECT g.genre
      FROM genres AS g
      JOIN genres_books AS g_b ON g_b.genre_id = g.id
      WHERE g_b.book_id = $1`,
      [bookID],
    );

    const genres = result.rows.map((row) => row.genre);

    return genres;
  }

  static async getNoGenreBooks(
    sort: string,
    offset: number,
    limit: number,
    ascDesc: string,
  ) {
    const result = await db.query<PostSavedMediaItem>(
      `SELECT b.*
      FROM books AS b
      LEFT JOIN genres_books AS gb
      ON gb.book_id = b.id
      WHERE gb.book_id IS NULL
      ORDER BY ${sort} ${ascDesc}
      LIMIT $1 OFFSET $2`,
      [limit, offset],
    );

    const totalRes = await db.query<{ count: string }>(
      `SELECT COUNT(*) 
          FROM books AS b
      LEFT JOIN genres_books AS gb
      ON gb.book_id = b.id
      WHERE gb.book_id IS NULL`,
    );

    return { books: result.rows, total: parseInt(totalRes.rows[0].count, 10) };
  }

  static async getBooksWithGenre(
    genre: string,
    sort: string,
    offset: number,
    limit: number,
    ascDesc: string,
  ) {
    const result = await db.query<PostSavedMediaItem>(
      `SELECT DISTINCT b.*
      FROM books AS b
      JOIN genres_books AS gb
      ON gb.book_id = b.id
      JOIN genres AS g
      ON g.id = gb.genre_id
      WHERE g.genre = $1
      ORDER BY ${sort} ${ascDesc}
      LIMIT $2 OFFSET $3
      `,
      [genre, limit, offset],
    );

    const totalRes = await db.query<{ count: string }>(
      `SELECT COUNT(*) 
          FROM books AS b
      JOIN genres_books AS gb
      ON gb.book_id = b.id
      JOIN genres AS g
      ON g.id = gb.genre_id
      WHERE g.genre = $1`,
      [genre],
    );
    return { books: result.rows, total: parseInt(totalRes.rows[0].count, 10) };
  }
}
