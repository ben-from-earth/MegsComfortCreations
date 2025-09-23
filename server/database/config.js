const path = require('path');
require('dotenv').config({
  path: path.resolve(__dirname, '../../.env'),
});

const username = process.env.PG_USERNAME;
const password = process.env.PG_PASSWORD;
const port = process.env.DB_PORT;

let DB_URI = `postgres://${username}:${password}@localhost:${port}`;

if (process.env.NODE_ENV === 'test') {
  DB_URI = `${DB_URI}/megscomfortcreations_test`;
} else {
  DB_URI = `${DB_URI}/megscomfortcreations`;
}

module.exports = DB_URI;
