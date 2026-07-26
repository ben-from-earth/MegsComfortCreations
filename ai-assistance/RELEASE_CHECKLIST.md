# Release Checklist (Local/Dev vs Production DB)

Use this checklist for every release so schema changes land on the right Neon branch.

## 1) Pre-Release (Feature Branch / Local)

- [ ] Generate migration SQL (if schema changed):
  - `npm run db:generate`
- [ ] Apply migration to Neon **dev** only:
  - `npm run db:migrate`
- [ ] (Optional) Reset Neon **dev** from snapshot before validating: `npm run db:restore` then `npm run db:migrate`
- [ ] Validate app behavior locally after migration.
- [ ] Review migration SQL in the PR before merge (especially destructive changes).

## 2) Merge and Deploy (happy path)

- [ ] Merge branch into `main`.
- [ ] Wait for Vercel **production** deployment to complete.
- [ ] In the deploy log, confirm Drizzle migrate ran against the expected prod host (hostname only), then `next build`.
- [ ] Smoke-test key production workflows.

Production migrations run automatically during the Vercel production build (`drizzle-migrate-ci.mjs`). You should **not** need a laptop prod migrate on a normal release.

## 3) Emergency: Manual Production Migration

Use only if auto-migrate failed or you must migrate without a deploy.

```bash
# Requires .env.production.local with prod DATABASE_URL / DATABASE_URL_UNPOOLED
npm run db:migrate:prod-emergency
```

Then fix the CI migrate path and redeploy.

## 4) Safety Rules (Always)

- [ ] Never run `db:migrate:prod-emergency` during regular development.
- [ ] Keep secrets only in env/secret stores (never in committed files).
- [ ] Preview deployments must not use production DB URLs or migrate production.
