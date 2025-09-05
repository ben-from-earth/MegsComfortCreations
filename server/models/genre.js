const db = require("../database/db");

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
      WHERE genre='${genre}'`
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
}

module.exports = Genre;
