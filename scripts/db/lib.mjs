import { config as loadDotenv } from 'dotenv';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { resolveMigrationDatabaseUrl } from '../../drizzle-migrate-shared.mjs';

const POSTGRES_CLIENT_HINT =
  'Install PostgreSQL client tools so pg_dump and psql are on PATH (https://www.postgresql.org/download/).';

/**
 * Neon dumps include the Drizzle migrations schema as well as public.
 * Wipe both before restore so `CREATE SCHEMA drizzle` in snapshot.sql succeeds.
 */
export const WIPE_TARGET_SCHEMAS_SQL = [
  'DROP SCHEMA IF EXISTS drizzle CASCADE',
  'DROP SCHEMA IF EXISTS public CASCADE',
  'CREATE SCHEMA public',
].join('; ');


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
  // override: true so the explicit env file wins over leftover shell exports
  // (e.g. a prior $env:DATABASE_URL from a manual dump). Without this,
  // snapshot:prod can silently dump the wrong Neon branch.
  const dotenvResult = loadDotenv({ path: envFilePath, override: true });
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
 * Resolve an executable on PATH without using a shell.
 * On Windows, prefer the absolute path from where.exe so spawn can run
 * without shell:true (Neon URLs contain `&`, which cmd.exe treats as a separator).
 * @param {string} commandName
 * @returns {string | null}
 */
export function resolveExecutable(commandName) {
  const probe =
    process.platform === 'win32'
      ? spawnSync('where.exe', [commandName], { encoding: 'utf8' })
      : spawnSync('which', [commandName], { encoding: 'utf8' });

  if (probe.status !== 0) {
    return null;
  }

  const firstMatch = probe.stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find(Boolean);

  return firstMatch ?? null;
}

/**
 * Spawn options for pg_dump/psql. Never use a shell — connection URLs may include
 * `&channel_binding=...` and other query params that break on Windows cmd.exe.
 * @param {string} command
 * @param {string[]} args
 * @param {NodeJS.ProcessEnv} [env]
 */
export function buildPostgresClientSpawn(command, args, env = process.env) {
  const executable = resolveExecutable(command) ?? command;
  return {
    command: executable,
    args,
    options: {
      stdio: 'inherit',
      env,
      // Critical: shell must stay false so `&` in DATABASE_URL is not interpreted.
      shell: false,
    },
  };
}

/**
 * @param {string} commandName
 */
export function assertPostgresClientAvailable(commandName) {
  if (!resolveExecutable(commandName)) {
    throw new Error(
      `Missing required command '${commandName}'. ${POSTGRES_CLIENT_HINT}`,
    );
  }
}

/**
 * @param {string} command
 * @param {string[]} args
 * @param {NodeJS.ProcessEnv} [env]
 */
export function runCommand(command, args, env = process.env) {
  const spawnConfig = buildPostgresClientSpawn(command, args, env);
  const result = spawnSync(
    spawnConfig.command,
    spawnConfig.args,
    spawnConfig.options,
  );

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
