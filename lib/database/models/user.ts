import db from '../db';

export class User {
  static async create({
    email,
    password,
    first_name,
    last_name,
  }: {
    email: string;
    password: string;
    first_name: string;
    last_name: string;
  }) {
    const result = await db.query(
      `INSERT INTO users (
            email,
            password_hash,
            first_name,
            last_name
      ) VALUES ($1, $2, $3, $4) 
      RETURNING 
            id,
            email,
            first_name,
            last_name`,
      [email, password, first_name, last_name],
    );
    return result.rows[0];
  }

  static async login(email: string) {
    const result = await db.query(
      `SELECT * FROM users
      WHERE email = $1`,
      [email],
    );
    return result.rows[0];
  }

  static async findOne(id: string) {
    const result = await db.query(
      `SELECT first_name, last_name, email FROM users
      WHERE id = $1`,
      [id],
    );
    return result.rows[0];
  }
}

export type UserType = {
  id: string;
  email: string;
  first_name: string;
  last_name: string;
};
