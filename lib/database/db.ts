import { Client } from "pg";
import DB_URI from "@/lib/database/config";

const db = new Client({
  connectionString: DB_URI,
});

async function initDb() {
  try {
    await db.connect();
    console.log("Successful connection to PostgreSQL database");
  } catch (err) {
    console.error("Connection to database failed:", err);
    process.exit(1);
  }
}

initDb();

export default db;
