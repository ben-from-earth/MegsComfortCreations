import fs from 'node:fs';
import path from 'node:path';
import {
  assertPostgresClientAvailable,
  loadEnvFile,
  resolveDatabaseTarget,
  runCommand,
} from './lib.mjs';

const envFilePath = process.argv[2];

if (!envFilePath) {
  console.error('Usage: node scripts/db/snapshot.mjs <env-file>');
  process.exit(1);
}

try {
  loadEnvFile(envFilePath);
  const target = resolveDatabaseTarget(process.env);
  assertPostgresClientAvailable('pg_dump');

  const snapshotPath = path.resolve('app/db/snapshot.sql');
  fs.mkdirSync(path.dirname(snapshotPath), { recursive: true });

  console.log(`Using ${target.source}`);
  console.log(`Snapshot source host: ${target.hostname}`);
  console.log(`Snapshot source database: ${target.databaseName}`);
  console.log(`Source env file: ${envFilePath}`);
  console.log(`Writing: ${snapshotPath}`);

  const status = runCommand('pg_dump', [
    target.url,
    '--no-owner',
    '--no-privileges',
    '--verbose',
    '-f',
    snapshotPath,
  ]);

  if (status !== 0) {
    console.error(`pg_dump failed with exit code ${status}`);
    process.exit(status);
  }

  console.log('Snapshot complete.');
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}
