const db = require("../db.cjs");

const convertToNull = (v) => (v === undefined ? null : v);

class Book {
  static async create(data) {
    const result = await db.query(
      `INSERT INTO books (
            title,
            author,
            page_count,
            pub_year,
            image_urls 
      ) VALUES ($1, $2, $3, $4, $5) 
      RETURNING 
            id,
            title,
            author,
            page_count,
            pub_year,
            image_urls`,
      [
        data.title,
        data.author,
        convertToNull(data.page_count),
        convertToNull(data.pub_year),
        convertToNull(data.image_urls),
      ]
    );
    return result.rows[0];
  }
}

module.exports = Book;
