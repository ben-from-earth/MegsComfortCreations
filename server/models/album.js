const db = require('../database/db');

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
      [data.title, data.image_urls]
    );
    return result.rows[0];
  }
  static async find(title) {
    const result = await db.query(
      `SELECT title, image_urls
     FROM albums
     WHERE title ILIKE '${title}'`
    );
    return result.rows;
  }
  static async edit(data) {
    const result = await db.query(
      `UPDATE albums
      SET title=$1, image_urls=$2
      WHERE id=$3
      RETURNING *`,
      [data.title, data.image_urls, data.id]
    );
    return result.rows[0];
  }
}

module.exports = Album;
