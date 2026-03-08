# Release Checklist (Local/Dev vs Production DB)

Use this checklist for every release so migrations are applied intentionally and only to the correct target.

## 1) Pre-Release (Feature Branch / Local)

- [ ] Confirm local app runtime target is dev/local branch:
  - `npm run db:check:local`
- [ ] Confirm local migration target is dev/local branch:
  - `npm run db:check:dev`
- [ ] Generate migration SQL (if schema changed):
  - `npm run db:generate`
- [ ] Apply migration to dev/local only:
  - `npm run db:migrate:dev`
- [ ] Validate app behavior locally after migration.

## 2) Merge and Deploy

- [ ] Merge branch into `main`.
- [ ] Wait for Vercel production deployment to complete.
- [ ] Confirm no production migration has run automatically (manual gate remains in place).

## 3) Production Migration Gate (Manual)

- [ ] Verify production DB target before running migration:
  - `npm run db:check:prod`
- [ ] Confirm output host/database matches intended production endpoint.
- [ ] Run production migration intentionally:
  - `npm run db:migrate:prod`

## 4) Post-Release Verification

- [ ] Smoke test key production workflows.
- [ ] Re-run production target check:
  - `npm run db:check:prod`
- [ ] Confirm migration state and data look correct in Neon production branch.

## 5) Safety Rules (Always)

- [ ] Never run production migration command during regular development.
- [ ] Keep secrets only in env/secret stores (never in committed files).
- [ ] If uncertain about target, run `db:check:*` command before any migration.
