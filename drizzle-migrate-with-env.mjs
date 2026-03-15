import { config as loadDotenv } from 'dotenv';
import { spawnSync } from 'node:child_process';

const envFilePath = process.argv[2];

if (!envFilePath) {
  console.error('Usage: node drizzle-migrate-with-env.mjs <env-file>');
  process.exit(1);
}

const dotenvResult = loadDotenv({ path: envFilePath });
if (dotenvResult.error) {
  console.error(`Unable to load env file: ${envFilePath}`);
  process.exit(1);
}

if (!process.env.DATABASE_URL) {
  console.error(`DATABASE_URL is missing in ${envFilePath}`);
  process.exit(1);
}

process.env.DRIZZLE_DATABASE_URL = process.env.DATABASE_URL;

const targetHost = new URL(process.env.DRIZZLE_DATABASE_URL).hostname;
console.log(`Running Drizzle migrate against host: ${targetHost}`);
console.log(`Source env file: ${envFilePath}`);

const command = process.platform === 'win32' ? 'npx drizzle-kit migrate' : 'npx';
const commandArgs = process.platform === 'win32' ? [] : ['drizzle-kit', 'migrate'];
const migrationRun = spawnSync(command, commandArgs, {
  stdio: 'inherit',
  env: process.env,
  shell: process.platform === 'win32',
});

if (migrationRun.error) {
  console.error('Failed to execute migration:', migrationRun.error.message);
  process.exit(1);
}

process.exit(migrationRun.status ?? 1);
