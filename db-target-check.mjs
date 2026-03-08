import { config as loadDotenv } from 'dotenv';
import pg from 'pg';

const envFilePath = process.argv[2] || '.env.local';
const dotenvResult = loadDotenv({ path: envFilePath });

if (dotenvResult.error) {
  console.error(`Unable to load env file: ${envFilePath}`);
  process.exit(1);
}

const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  console.error(`DATABASE_URL is missing in ${envFilePath}`);
  process.exit(1);
}

const parsedUrl = new URL(databaseUrl);
const databaseHost = parsedUrl.hostname;
const databaseName = parsedUrl.pathname.replace(/^\//, '') || '(unknown)';

const { Client } = pg;
const client = new Client({ connectionString: databaseUrl });

const run = async () => {
  await client.connect();

  const tableCountQuery = `
    select
      (select count(*) from books) as books_count,
      (select count(*) from movies) as movies_count,
      (select count(*) from video_games) as video_games_count,
      (select count(*) from albums) as albums_count
  `;

  const [{ current_database: currentDatabase }] = (
    await client.query('select current_database()')
  ).rows;
  const [tableCounts] = (await client.query(tableCountQuery)).rows;

  console.log('DB target summary');
  console.log(`- env file: ${envFilePath}`);
  console.log(`- host: ${databaseHost}`);
  console.log(`- database from URL: ${databaseName}`);
  console.log(`- current_database(): ${currentDatabase}`);
  console.log('- table counts:');
  console.log(`  - books: ${tableCounts.books_count}`);
  console.log(`  - movies: ${tableCounts.movies_count}`);
  console.log(`  - video_games: ${tableCounts.video_games_count}`);
  console.log(`  - albums: ${tableCounts.albums_count}`);
};

run()
  .catch((error) => {
    console.error('Failed to inspect DB target:', error.message);
    process.exitCode = 1;
  })
  .finally(async () => {
    await client.end();
  });
