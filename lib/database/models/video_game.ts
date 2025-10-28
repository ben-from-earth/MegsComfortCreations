// database import
import db from '@/lib/database/db';

// interfaces and types
import {
  MediaType,
  postSavedMediaItem,
  presavedMediaItem,
} from '@/lib/interfaces/globalInterfaces';

export default class Video_Game {
  static async create(data: presavedMediaItem) {
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
      [data.title, data.image_urls, data.spine_color],
    );
    return result.rows[0];
  }

  static async find(title: string) {
    const result = await db.query<postSavedMediaItem>(
      `SELECT * 
     FROM video_games
     WHERE title ILIKE $1 || '%'`,
      [title],
    );
    return result.rows;
  }

  static async findSome(
    type: MediaType,
    offset: number,
    limit: number,
    sort: string,
  ) {
    const result = await db.query<postSavedMediaItem>(
      `SELECT * 
      FROM ${type + 's'}
      ORDER BY ${sort}
      LIMIT $1 OFFSET $2`,
      [limit, offset],
    );
    return result.rows;
  }

  static async edit(data: postSavedMediaItem) {
    const result = await db.query(
      `UPDATE video_games
      SET title=$1, image_urls=$2, spine_color=$3
      WHERE id=$4
      RETURNING *`,
      [data.title, data.image_urls, data.spine_color, data.id],
    );
    return result.rows[0];
  }
}
