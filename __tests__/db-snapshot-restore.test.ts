/**
 * Safety-guard coverage for Neon snapshot/restore helpers.
 * Does not connect to a real database.
 */

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

type DbLibModule = {
  assertDevRestoreTarget: (envFilePath: string) => void;
  loadEnvFile: (envFilePath: string) => void;
  WIPE_TARGET_SCHEMAS_SQL: string;
  resolveDatabaseTarget: (env: NodeJS.ProcessEnv) => {
    url: string;
    source: string;
    hostname: string;
    databaseName: string;
  };
  buildPostgresClientSpawn: (
    command: string,
    args: string[],
    env?: NodeJS.ProcessEnv,
  ) => {
    command: string;
    args: string[];
    options: { shell: boolean };
  };
};

async function loadDbLib(): Promise<DbLibModule> {
  return import('../scripts/db/lib.mjs') as Promise<DbLibModule>;
}

describe('scripts/db restore safety', () => {
  const originalDatabaseUrl = process.env.DATABASE_URL;
  const originalUnpooledUrl = process.env.DATABASE_URL_UNPOOLED;
  const originalDrizzleUrl = process.env.DRIZZLE_DATABASE_URL;

  afterEach(() => {
    if (originalDatabaseUrl === undefined) {
      delete process.env.DATABASE_URL;
    } else {
      process.env.DATABASE_URL = originalDatabaseUrl;
    }
    if (originalUnpooledUrl === undefined) {
      delete process.env.DATABASE_URL_UNPOOLED;
    } else {
      process.env.DATABASE_URL_UNPOOLED = originalUnpooledUrl;
    }
    if (originalDrizzleUrl === undefined) {
      delete process.env.DRIZZLE_DATABASE_URL;
    } else {
      process.env.DRIZZLE_DATABASE_URL = originalDrizzleUrl;
    }
  });

  it('refuses production env file paths', async () => {
    const { assertDevRestoreTarget } = await loadDbLib();

    expect(() =>
      assertDevRestoreTarget('.env.production.local'),
    ).toThrow(/production/i);
    expect(() =>
      assertDevRestoreTarget('env/.env.production'),
    ).toThrow(/production/i);
    expect(() => assertDevRestoreTarget('prod.env')).toThrow(/production/i);
  });

  it('allows development env file paths', async () => {
    const { assertDevRestoreTarget } = await loadDbLib();

    expect(() =>
      assertDevRestoreTarget('.env.development.local'),
    ).not.toThrow();
    expect(() =>
      assertDevRestoreTarget('C:\\repo\\.env.development.local'),
    ).not.toThrow();
  });

  it('refuses ambiguous env files that are not development targets', async () => {
    const { assertDevRestoreTarget } = await loadDbLib();

    expect(() => assertDevRestoreTarget('.env.local')).toThrow(/development/i);
    expect(() => assertDevRestoreTarget('.env')).toThrow(/development/i);
  });

  it('resolves database target preferring unpooled URL', async () => {
    const { resolveDatabaseTarget } = await loadDbLib();

    const target = resolveDatabaseTarget({
      DATABASE_URL_UNPOOLED: 'postgres://user:secret@ep-dev.neon.tech/neondb',
      DATABASE_URL: 'postgres://user:secret@ep-dev-pooler.neon.tech/neondb',
    });

    expect(target.source).toBe('DATABASE_URL_UNPOOLED');
    expect(target.hostname).toBe('ep-dev.neon.tech');
    expect(target.databaseName).toBe('neondb');
  });

  it('wipes drizzle and public schemas before restore', async () => {
    const { WIPE_TARGET_SCHEMAS_SQL } = await loadDbLib();
    expect(WIPE_TARGET_SCHEMAS_SQL).toMatch(/DROP SCHEMA IF EXISTS drizzle CASCADE/i);
    expect(WIPE_TARGET_SCHEMAS_SQL).toMatch(/DROP SCHEMA IF EXISTS public CASCADE/i);
    expect(WIPE_TARGET_SCHEMAS_SQL).toMatch(/CREATE SCHEMA public/i);
  });

  it('spawns pg clients without a shell so URL query & is preserved', async () => {
    const { buildPostgresClientSpawn } = await loadDbLib();
    const neonUrl =
      'postgresql://user:pass@ep-example.neon.tech/neondb?sslmode=require&channel_binding=require';

    const spawnConfig = buildPostgresClientSpawn('pg_dump', [
      neonUrl,
      '--no-owner',
      '-f',
      'app/db/snapshot.sql',
    ]);

    expect(spawnConfig.options.shell).toBe(false);
    expect(spawnConfig.args[0]).toContain('channel_binding=require');
    expect(spawnConfig.args[0]).toContain('&');
  });

  it('overrides leftover shell DATABASE_URL when loading an env file', async () => {
    const { loadEnvFile, resolveDatabaseTarget } = await loadDbLib();
    const tempEnvPath = path.join(
      os.tmpdir(),
      `mcc-snapshot-env-${Date.now()}.env`,
    );

    fs.writeFileSync(
      tempEnvPath,
      'DATABASE_URL_UNPOOLED=postgres://user:pass@ep-from-file.neon.tech/neondb\n',
      'utf8',
    );

    try {
      process.env.DATABASE_URL_UNPOOLED =
        'postgres://user:pass@ep-from-shell.neon.tech/neondb';
      loadEnvFile(tempEnvPath);
      const target = resolveDatabaseTarget(process.env);
      expect(target.hostname).toBe('ep-from-file.neon.tech');
    } finally {
      fs.unlinkSync(tempEnvPath);
    }
  });
});
