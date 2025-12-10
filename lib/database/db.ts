import { Client } from 'pg';
const connectionString = process.env.DB_CONNECTION_STRING as string;

export const DB_URI: string = connectionString;

const db = new Client({
  connectionString: DB_URI,
});

async function initDb() {
  try {
    await db.connect();
    console.log('Successful connection to PostgreSQL database');
  } catch (err) {
    console.error('Connection to database failed:', err);
    process.exit(1);
  }
}

initDb();

export default db;
