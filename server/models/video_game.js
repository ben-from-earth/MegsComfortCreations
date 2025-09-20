const db = require('../database/db');

class Video_Game {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO video_games (
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
     FROM video_games
     WHERE title ILIKE '${title}'`
    );
    return result.rows;
  }
}

module.exports = Video_Game;
