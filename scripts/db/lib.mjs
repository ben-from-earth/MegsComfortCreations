import { config as loadDotenv } from 'dotenv';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { resolveMigrationDatabaseUrl } from '../../drizzle-migrate-shared.mjs';

const POSTGRES_CLIENT_HINT =
  'Install PostgreSQL client tools so pg_dump and psql are on PATH (https://www.postgresql.org/download/).';

/**
 * Refuse restore unless the env file is clearly a development target.
 * Blocks any path that looks like production (e.g. .env.production.local).
 * @param {string} envFilePath
 */
export function assertDevRestoreTarget(envFilePath) {
  const normalizedPath = envFilePath.replaceAll('\\', '/').toLowerCase();
  const baseName = path.basename(normalizedPath);

  // Prod markers — never allow restore against these files.
  if (
    normalizedPath.includes('.env.production') ||
    normalizedPath.includes('production.local') ||
    baseName.includes('production') ||
    baseName.includes('prod')
  ) {
    throw new Error(
      `Refusing restore: env file looks like production (${envFilePath}). Restore is allowed for development env files only.`,
    );
  }

  // Require an explicit development marker so random env files cannot be wiped.
  if (
    !normalizedPath.includes('.env.development') &&
    !normalizedPath.includes('development.local') &&
    !baseName.includes('development') &&
    !baseName.includes('dev')
  ) {
    throw new Error(
      `Refusing restore: env file is not a development target (${envFilePath}). Use .env.development.local.`,
    );
  }
}

/**
 * @param {string} envFilePath
 */
export function loadEnvFile(envFilePath) {
  const dotenvResult = loadDotenv({ path: envFilePath });
  if (dotenvResult.error) {
    throw new Error(`Unable to load env file: ${envFilePath}`);
  }
}

/**
 * @param {NodeJS.ProcessEnv} env
 */
export function resolveDatabaseTarget(env) {
  const resolved = resolveMigrationDatabaseUrl(env);
  if (!resolved) {
    throw new Error(
      'DATABASE_URL_UNPOOLED or DATABASE_URL is missing in the loaded env file.',
    );
  }

  let parsedUrl;
  try {
    parsedUrl = new URL(resolved.url);
  } catch {
    throw new Error('Resolved database URL is invalid.');
  }

  const databaseName = parsedUrl.pathname.replace(/^\//, '') || '(unknown)';

  return {
    url: resolved.url,
    source: resolved.source,
    hostname: parsedUrl.hostname,
    databaseName,
  };
}

/**
 * @param {string} commandName
 */
export function assertPostgresClientAvailable(commandName) {
  const probe =
    process.platform === 'win32'
      ? spawnSync(`where ${commandName}`, {
          encoding: 'utf8',
          shell: true,
        })
      : spawnSync('which', [commandName], { encoding: 'utf8' });

  if (probe.status !== 0) {
    throw new Error(`Missing required command '${commandName}'. ${POSTGRES_CLIENT_HINT}`);
  }
}

/**
 * @param {string} command
 * @param {string[]} args
 * @param {NodeJS.ProcessEnv} [env]
 */
export function runCommand(command, args, env = process.env) {
  const useShell = process.platform === 'win32';
  const result = spawnSync(command, args, {
    stdio: 'inherit',
    env,
    shell: useShell,
  });

  if (result.error) {
    if (result.error.code === 'ENOENT') {
      throw new Error(`Failed to run '${command}'. ${POSTGRES_CLIENT_HINT}`);
    }
    throw new Error(`Failed to run '${command}': ${result.error.message}`);
  }

  return result.status ?? 1;
}

/**
 * @param {string} message
 * @param {NodeJS.ReadStream} [input]
 * @param {NodeJS.WriteStream} [output]
 * @returns {Promise<string>}
 */
export async function promptForConfirmation(
  message,
  input = process.stdin,
  output = process.stdout,
) {
  const { createInterface } = await import('node:readline/promises');
  const rl = createInterface({ input, output });
  try {
    const answer = await rl.question(message);
    return answer.trim().toLowerCase();
  } finally {
    rl.close();
  }
}
