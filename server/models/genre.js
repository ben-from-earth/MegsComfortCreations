const db = require('../database/db');

class Genre {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO genres (
            genre
      ) VALUES ($1) 
      RETURNING 
            id,
            genre`,
      data.genre
    );
    return result.rows[0];
  }
  static async getAllGenres() {
    const result = await db.query(`SELECT genre FROM genres`);
    return result.rows;
  }

  static async link(genre, bookID) {
    const genreFind = await db.query(
      `SELECT id FROM genres
      WHERE genre=$1`,
      [genre]
    );

    const genreID = genreFind.rows[0].id;
    const genreLinkResult = await db.query(
      `INSERT INTO genres_books (
        book_id, genre_id
      ) VALUES ($1, $2)`,
      [bookID, genreID]
    );
    return genreLinkResult;
  }

  static async unlink(bookID) {
    const genreUnLinkResult = await db.query(
      `DELETE FROM genres_books
      WHERE book_id = $1`,
      [bookID]
    );
    return genreUnLinkResult;
  }

  static async getFromBook(bookID) {
    const result = await db.query(
      `SELECT g.genre
      FROM genres AS g
      JOIN genres_books AS g_b ON g_b.genre_id::uuid = g.id
      WHERE g_b.book_id = $1`,
      [bookID]
    );
    const genres = result.rows.map((row) => row.genre);

    return genres;
  }

  static async getNoGenreBooks(sort, offset, limit) {
    const result = await db.query(
      `SELECT b.*
      FROM books AS b
      LEFT JOIN genres_books AS gb
      ON gb.book_id = b.id::text
      WHERE gb.book_id IS NULL
      ORDER BY ${sort}
      LIMIT $1 OFFSET $2`,
      [limit, offset]
    );

    const totalRes = await db.query(
      `SELECT COUNT(*) 
          FROM books AS b
      LEFT JOIN genres_books AS gb
      ON gb.book_id = b.id::text
      WHERE gb.book_id IS NULL`
    );

    return { books: result.rows, total: parseInt(totalRes.rows[0].count, 10) };
  }

  static async getBooksWithGenre(genre, sort, offset, limit) {
    const result = await db.query(
      `SELECT DISTINCT b.*
      FROM books AS b
      JOIN genres_books AS gb
      ON gb.book_id = b.id::text
      JOIN genres AS g
      ON g.id::text = gb.genre_id
      WHERE g.genre = $1
      ORDER BY ${sort}
      LIMIT $2 OFFSET $3
      `,
      [genre, limit, offset]
    );

    const totalRes = await db.query(
      `SELECT COUNT(*) 
          FROM books AS b
      JOIN genres_books AS gb
      ON gb.book_id = b.id::text
      JOIN genres AS g
      ON g.id::text = gb.genre_id
      WHERE g.genre = $1`,
      [genre]
    );
    return { books: result.rows, total: parseInt(totalRes.rows[0].count, 10) };
  }
}

module.exports = Genre;
