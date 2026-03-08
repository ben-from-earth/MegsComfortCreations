# Neon Database Targeting Workflow

This workflow keeps development and production database changes separated when using Neon branches with Vercel.

## One-time setup

Create these local files (not committed):
- `.env.local` -> local app runtime target (usually Neon dev branch while developing)
- `.env.development.local` -> migration target for dev branch
- `.env.production.local` -> migration target for prod branch

Each file needs:
- `DATABASE_URL=postgresql://...`

## Verify which database you are targeting

Use the check scripts:

```bash
npm run db:check:local
npm run db:check:dev
npm run db:check:prod
```

Each script prints:
- URL host (endpoint fingerprint)
- database name
- `current_database()`
- row counts from key tables (`books`, `movies`, `video_games`, `albums`)

Use this to match what you see in Neon SQL editor for the same branch.

## Safe migration flow

1. Generate SQL from schema changes:
   - `npm run db:generate`
2. Apply migration to dev branch only:
   - `npm run db:migrate:dev`
3. Validate app and data in local/preview using dev branch URL.
4. Release app code.
5. Apply the exact migration to prod branch intentionally:
   - `npm run db:migrate:prod`
6. Follow the full release checklist:
   - `ai-assistance/RELEASE_CHECKLIST.md`

## Vercel mapping recommendation

- **Preview environment** `DATABASE_URL` -> Neon dev branch URL
- **Production environment** `DATABASE_URL` -> Neon prod branch URL

## Preflight target check (before every migration)

1. Run the matching check command:
   - local runtime: `npm run db:check:local`
   - dev migration: `npm run db:check:dev`
   - prod migration: `npm run db:check:prod`
2. Confirm the printed host/database is the intended target.
3. Only then run the corresponding migration command.
