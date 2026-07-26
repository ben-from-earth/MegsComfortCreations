import fs from 'node:fs';
import path from 'node:path';
import {
  assertDevRestoreTarget,
  assertPostgresClientAvailable,
  loadEnvFile,
  promptForConfirmation,
  resolveDatabaseTarget,
  runCommand,
} from './lib.mjs';

const args = process.argv.slice(2);
const forceFlag = args.includes('-f') || args.includes('--force');
const envFilePath = args.find((arg) => !arg.startsWith('-'));

if (!envFilePath) {
  console.error('Usage: node scripts/db/restore.mjs <env-file> [--force|-f]');
  process.exit(1);
}

try {
  assertDevRestoreTarget(envFilePath);
  loadEnvFile(envFilePath);
  const target = resolveDatabaseTarget(process.env);
  assertPostgresClientAvailable('psql');

  const snapshotPath = path.resolve('app/db/snapshot.sql');
  if (!fs.existsSync(snapshotPath)) {
    console.error(
      `Missing snapshot file: ${snapshotPath}. Run npm run db:snapshot or npm run db:snapshot:prod first.`,
    );
    process.exit(1);
  }

  console.log(`Using ${target.source}`);
  console.log(`Restore target host: ${target.hostname}`);
  console.log(`Restore target database: ${target.databaseName}`);
  console.log(`Source env file: ${envFilePath}`);
  console.log(`Snapshot file: ${snapshotPath}`);

  if (!forceFlag) {
    const answer = await promptForConfirmation(
      `This will WIPE all data on host ${target.hostname} / database ${target.databaseName}. Type yes to continue: `,
    );
    if (answer !== 'yes') {
      console.log('Restore cancelled.');
      process.exit(0);
    }
  } else {
    console.log('Force flag set; skipping confirmation prompt.');
  }

  const wipeStatus = runCommand('psql', [
    target.url,
    '-v',
    'ON_ERROR_STOP=1',
    '-c',
    'DROP SCHEMA public CASCADE; CREATE SCHEMA public;',
  ]);
  if (wipeStatus !== 0) {
    console.error(`Schema wipe failed with exit code ${wipeStatus}`);
    process.exit(wipeStatus);
  }

  const restoreStatus = runCommand('psql', [
    target.url,
    '-v',
    'ON_ERROR_STOP=1',
    '-f',
    snapshotPath,
  ]);
  if (restoreStatus !== 0) {
    console.error(`Snapshot restore failed with exit code ${restoreStatus}`);
    process.exit(restoreStatus);
  }

  console.log('Restore complete.');
  console.log('Next step: npm run db:migrate');
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}
