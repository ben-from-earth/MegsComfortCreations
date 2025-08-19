const db = require("../db.cjs");

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
}

module.exports = Genre;
