import { config as loadDotenv } from 'dotenv';
import { spawnSync } from 'node:child_process';
import { resolveMigrationDatabaseUrl } from './drizzle-migrate-shared.mjs';

const envFilePath = process.argv[2];

if (!envFilePath) {
  console.error('Usage: node drizzle-migrate-with-env.mjs <env-file>');
  process.exit(1);
}

// override: true so the chosen env file wins over leftover shell DATABASE_URL* exports
const dotenvResult = loadDotenv({ path: envFilePath, override: true });
if (dotenvResult.error) {
  console.error(`Unable to load env file: ${envFilePath}`);
  process.exit(1);
}

const resolved = resolveMigrationDatabaseUrl(process.env);
if (!resolved) {
  console.error(
    `DATABASE_URL_UNPOOLED or DATABASE_URL is missing in ${envFilePath}`,
  );
  process.exit(1);
}

process.env.DRIZZLE_DATABASE_URL = resolved.url;

let targetHost;
try {
  targetHost = new URL(resolved.url).hostname;
} catch {
  console.error(`Resolved migration database URL in ${envFilePath} is invalid.`);
  process.exit(1);
}

console.log(`Using ${resolved.source}`);
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
