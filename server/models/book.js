const db = require("../database/db");

const convertToNull = (v) => (v === undefined ? null : v);

class Book {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO books (
            title,
            author,
            page_count,
            pub_year,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3, $4, $5, $6) 
      RETURNING 
            id,
            title,
            author,
            page_count,
            pub_year,
            image_urls,
            spine_color`,
      [
        data.title,
        data.author,
        convertToNull(data.page_count),
        convertToNull(data.pub_year),
        data.image_urls,
        data.spine_color,
      ]
    );
    return result.rows[0];
  }

  static async find(title) {
    const result = await db.query(
      `SELECT * 
     FROM books
     WHERE title ILIKE $1`,
      [title]
    );
    return result.rows;
  }

  static async findSome(type, offset, limit, sort) {
    const result = await db.query(
      `SELECT * 
      FROM ${type + "s"}
      ORDER BY ${sort}
      LIMIT $1 OFFSET $2`,
      [limit, offset]
    );
    return result.rows;
  }
}

module.exports = Book;
