const db = require('../database/db');

class Movie {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO movies (
            title,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3) 
      RETURNING 
            id,
            title,
            image_urls,
            spine_color`,
      [data.title, data.image_urls, data.spine_color]
    );
    return result.rows[0];
  }
  static async find(title) {
    const result = await db.query(
      `SELECT title, image_urls, spine_color 
     FROM movies
     WHERE title ILIKE '${title}'`
    );
    return result.rows;
  }
  static async edit(data) {
    const result = await db.query(
      `UPDATE movies
      SET title=$1, image_urls=$2, spine_color=$3
      WHERE id=$4
      RETURNING *`,
      [data.title, data.image_urls, data.spine_color, data.id]
    );
    return result.rows[0];
  }
}

module.exports = Movie;
