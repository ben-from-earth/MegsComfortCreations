import { spawnSync } from 'node:child_process';
import {
  resolveMigrationDatabaseUrl,
  shouldRunMigrateOnBuild,
} from './drizzle-migrate-shared.mjs';

if (!shouldRunMigrateOnBuild(process.env)) {
  console.log(
    'Skipping Drizzle migrate (not production build; set MIGRATE_ON_BUILD=true to force)',
  );
  process.exit(0);
}

const resolved = resolveMigrationDatabaseUrl(process.env);
if (!resolved) {
  console.error(
    'Missing database URL for migrate-on-build. Set DATABASE_URL_UNPOOLED (preferred) or DATABASE_URL.',
  );
  process.exit(1);
}

process.env.DRIZZLE_DATABASE_URL = resolved.url;

let targetHost;
try {
  targetHost = new URL(resolved.url).hostname;
} catch {
  console.error('Resolved migration database URL is invalid.');
  process.exit(1);
}

console.log(`Using ${resolved.source}`);
console.log(`Running Drizzle migrate against host: ${targetHost}`);

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
