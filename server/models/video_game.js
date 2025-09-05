const db = require("../database/db");

const convertToNull = (v) => (v === undefined ? null : v);

class Video_Game {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO video_games (
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

  static async find(title) {
    const result = await db.query(
      `SELECT title, image_urls 
     FROM video_games
     WHERE title ILIKE '${title}'`
    );
    return result.rows;
  }
}

module.exports = Video_Game;
