const db = require("../db.cjs");

const convertToNull = (v) => (v === undefined ? null : v);

class Movie {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO movies (
            title,
            image_urls 
      ) VALUES ($1, $2) 
      RETURNING 
            id,
            title,
            image_urls`,
      [data.title, convertToNull(data.image_urls)]
    );
    return result.rows[0];
  }
}

module.exports = Movie;
