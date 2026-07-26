/**
 * Shared helpers for Drizzle migrate entrypoints.
 * Prefer unpooled/direct Neon URL for migrations; app runtime can stay pooled.
 */

/**
 * @param {NodeJS.ProcessEnv} env
 * @returns {{ url: string, source: 'DRIZZLE_DATABASE_URL' | 'DATABASE_URL_UNPOOLED' | 'DATABASE_URL' } | null}
 */
export function resolveMigrationDatabaseUrl(env) {
  if (env.DRIZZLE_DATABASE_URL) {
    return { url: env.DRIZZLE_DATABASE_URL, source: 'DRIZZLE_DATABASE_URL' };
  }
  if (env.DATABASE_URL_UNPOOLED) {
    return {
      url: env.DATABASE_URL_UNPOOLED,
      source: 'DATABASE_URL_UNPOOLED',
    };
  }
  if (env.DATABASE_URL) {
    return { url: env.DATABASE_URL, source: 'DATABASE_URL' };
  }
  return null;
}

/**
 * Production Vercel builds always migrate. Local/preview skip unless explicitly forced.
 * @param {NodeJS.ProcessEnv} env
 */
export function shouldRunMigrateOnBuild(env) {
  return env.VERCEL_ENV === 'production' || env.MIGRATE_ON_BUILD === 'true';
}
