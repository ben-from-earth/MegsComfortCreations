// database import
import db from '@/lib/database/db';

// interfaces and types
import {
  MediaType,
  postSavedMediaItem,
  presavedMediaItem,
} from '@/lib/interfaces/globalInterfaces';

const convertToNull = (v: number | undefined) => (v === undefined ? null : v);

export default class Book {
  static async create(data: presavedMediaItem) {
    const result = await db.query<postSavedMediaItem>(
      `INSERT INTO books (
            title,
            author,
            page_count,
            pub_year,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3, $4, $5, $6) 
      RETURNING *`,
      [
        data.title,
        data.author,
        convertToNull(data.page_count),
        convertToNull(data.pub_year),
        data.image_urls,
        data.spine_color,
      ],
    );
    return result.rows[0];
  }

  static async find(title: string) {
    const result = await db.query<postSavedMediaItem>(
      `SELECT * 
     FROM books
     WHERE title ILIKE $1 || '%'
     ORDER BY title`,
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
    const result = await db.query<postSavedMediaItem>(
      `UPDATE books
      SET title = $1, author = $2, page_count = $3, pub_year = $4, image_urls = $5, spine_color = $6
      WHERE id=$7
      RETURNING *`,
      [
        data.title,
        data.author,
        data.page_count,
        data.pub_year,
        data.image_urls,
        data.spine_color,
        data.id,
      ],
    );
    return result.rows[0];
  }
}
