const db = require("../db.cjs");

const convertToNull = (v) => (v === undefined ? null : v);

class Album {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO albums (
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

module.exports = Album;
