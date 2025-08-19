// Database config
require("dotenv").config({
  path: require("path").resolve(__dirname, "../.env"),
});

const { Pool } = require("pg");
let db = new Pool({
  host: "localhost",
  port: 5432,
  database: "megscomfortcreations",
  user: "ben-from-earth",
  password: process.env.PG_PASSWORD,
});

db.connect()
  .then((client) => {
    console.log("Connected to Postgress DB");
    // client.release();
  })
  .catch((err) => console.error("Database Connection error", err.stack));

module.exports = db;
