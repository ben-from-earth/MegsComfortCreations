# Plan: Neon Snapshot + Restore (Dev Reset Loop)

## Goal

Add a local SquareRigger-style workflow for Neon:

1. **`db:snapshot`** — dump a known-good Neon DB into `app/db/snapshot.sql`
2. **`db:restore`** — wipe the **dev** target DB and reload that file
3. **`db:migrate`** — apply any migrations newer than the snapshot watermark

Use this for a repeatable “reset my Neon dev DB to a baseline, then catch up migrations” loop.

## Target End State

```text
npm run db:snapshot          # dump chosen source → app/db/snapshot.sql
npm run db:restore           # wipe Neon DEV + load snapshot.sql (confirm first)
npm run db:migrate           # apply migrations newer than the dump
```

Safety rules (must be enforced in code):

- Restore targets **dev only** (`.env.development.local`). Never prod by default.
- Prefer **direct / unpooled** URLs for dump/restore (`DATABASE_URL_UNPOOLED`, then `DATABASE_URL`).
- Snapshot file is gitignored (may contain real data / PII).
- Scripts log **hostname + database name only** (never password / full URL).
- Destructive restore requires interactive confirmation unless `-f` / `--force`.

Neon note: do **not** use `DROP DATABASE` / `CREATE DATABASE`. Restore = wipe `public` schema, then `psql -f`.

## Already Done (do not redo)

- [x] `npm run db:migrate` → `.env.development.local` via `drizzle-migrate-with-env.mjs`
- [x] `npm run db:migrate:prod-emergency` → laptop prod escape hatch only
- [x] Shared URL resolver in `drizzle-migrate-shared.mjs`:
  - `DRIZZLE_DATABASE_URL` → `DATABASE_URL_UNPOOLED` → `DATABASE_URL`
- [x] Prod auto-migrate on Vercel: see `ai-assistance/PLAN_DB_MIGRATE_VERCEL.md` (separate work)
- [x] Human preferred baseline source: **B — prod explicit** (`npm run db:snapshot:prod`)
- [x] Snapshot/restore scripts, gitignore, docs, safety tests (implementation complete; human commits)

---

## Human Input Required

Do these **before** asking the agent to implement (or in parallel while the agent writes scripts). The agent cannot install Postgres tools or put secrets in your env files.

### H1) Install Postgres client tools

You need `pg_dump` and `psql` on PATH.

Windows options:

- Install [PostgreSQL](https://www.postgresql.org/download/windows/) (client tools are enough), or
- `winget install PostgreSQL.PostgreSQL` and ensure the `bin` directory is on PATH

Verify in PowerShell:

```powershell
pg_dump --version
psql --version
```

### H2) Confirm local Neon env files (gitignored)

| File | Role |
|------|------|
| `.env.development.local` | Neon **dev** — restore target; default snapshot source |
| `.env.production.local` | Neon **prod** — optional snapshot **source** only; never restore target |
| `.env.local` | App runtime (pooled URL fine) — not used by snapshot/restore |

For dump/restore, prefer a **direct** (non-`-pooler`) URL:

1. Neon Console → project → **dev** branch → Connection details.
2. Copy the **direct** connection string.
3. Put it in `.env.development.local` as either:
   - `DATABASE_URL_UNPOOLED=...` (preferred), and keep pooled `DATABASE_URL` for the app if you use the same file, **or**
   - `DATABASE_URL=...` set to the direct URL for this migration/snapshot env file.

If you want `db:snapshot:prod` later, do the same for `.env.production.local` with the **prod** branch direct URL. Never use that file as a restore target.

### H3) Decide snapshot source policy

Pick one default and stick to it:

| Choice | When |
|--------|------|
| **A — Neon dev (recommended default)** | Baseline for day-to-day reset after a good seed |
| **B — Prod dump (explicit only)** | Capture real data for debugging — careful with PII |
| **C — Throwaway Neon branch from prod** | Safest prod-like data; snapshot from that branch URL |

Agent will implement:

- `db:snapshot` → defaults to `.env.development.local`
- `db:snapshot:prod` → `.env.production.local` (explicit; read-only dump)
- `db:restore` → **only** `.env.development.local` (refuse production env files)

### H4) One-time manual smoke (optional but recommended)

Before relying on scripts, optionally prove dump/restore works once (PowerShell from repo root):

```powershell
# Prefer DATABASE_URL_UNPOOLED if present; else DATABASE_URL
$line = Get-Content .env.development.local |
  Where-Object { $_ -match '^\s*(DATABASE_URL_UNPOOLED|DATABASE_URL)\s*=' } |
  Select-Object -First 1
$env:DATABASE_URL = ($line -split '=', 2)[1].Trim()
([Uri]$env:DATABASE_URL).Host   # sanity: should be your Neon DEV host (no -pooler preferred)

pg_dump $env:DATABASE_URL --no-owner --no-privileges --verbose -f app/db/snapshot.sql
```

Do **not** run the wipe/restore half against prod. If you want to test restore, only against **dev**.

### Hand-off back to agent

Paste this filled block when H1–H3 are done (H4 optional):

```text
Human setup complete:
- pg_dump / psql on PATH: yes/no
- .env.development.local has direct URL (DATABASE_URL_UNPOOLED or DATABASE_URL): yes/no
- .env.production.local ready for optional prod snapshot: yes/no/skip
- Default snapshot source: A (dev) / B (prod explicit) / C (throwaway branch)
- Manual dump smoke tested (optional): yes/no/skip
- Ready for agent to implement: yes/no
```

Do **not** paste secret connection strings here.

---

## Agent Implementation Steps

Execute in order after human hand-off. Stop and ask if something conflicts with repo state.

### A0) Preconditions

1. Read `ai-assistance/LEARNINGS.md` and this plan.
2. Confirm human hand-off block is complete (at least H1–H3).
3. Reuse `resolveMigrationDatabaseUrl` from `drizzle-migrate-shared.mjs` (do not invent a second resolution order).
4. Confirm `db:migrate` already exists; do not rename it.

### A1) Add shared DB script helpers

Create `scripts/db/lib.mjs` (or equivalent) that:

1. Loads a dotenv file path argument.
2. Resolves URL via `resolveMigrationDatabaseUrl(process.env)`.
3. Exposes helpers to:
   - Parse hostname + database name for logging
   - Detect missing `pg_dump` / `psql` with a clear install hint
   - Spawn shell commands cross-platform (same win32 care as migrate scripts)
4. Exposes `assertDevRestoreTarget(envFilePath)` that **refuses** restore when:
   - Path contains `production` / `.env.production`, **or**
   - Any other explicit prod marker the agent documents in code comments

### A2) Create `scripts/db/snapshot.mjs`

Behavior:

1. Usage: `node scripts/db/snapshot.mjs <env-file>`
2. Load env → resolve URL (prefer unpooled).
3. Log hostname + db name + which env source key was used.
4. Run:

   ```bash
   pg_dump "$URL" --no-owner --no-privileges --verbose -f app/db/snapshot.sql
   ```

5. Ensure `app/db/` exists; overwrite `app/db/snapshot.sql`.
6. Exit non-zero if `pg_dump` missing or dump fails.
7. Never print the full URL / password.

### A3) Create `scripts/db/restore.mjs`

Behavior:

1. Usage: `node scripts/db/restore.mjs <env-file> [--force|-f]`
2. Call `assertDevRestoreTarget` — fail hard if not a dev env file.
3. Require `app/db/snapshot.sql` exists; clear error if missing (`run db:snapshot first`).
4. Log hostname + db name.
5. Unless `-f` / `--force`, prompt on stdin:

   `This will WIPE all data on host X / database Y. Type yes to continue:`

   Exit 0 without changes if not confirmed.
6. Wipe schema (Neon-friendly):

   ```sql
   DROP SCHEMA public CASCADE;
   CREATE SCHEMA public;
   ```

7. Restore:

   ```bash
   psql "$URL" -v ON_ERROR_STOP=1 -f app/db/snapshot.sql
   ```

8. Print next step: `npm run db:migrate`
9. Do **not** implement `DROP DATABASE`.

### A4) Wire `package.json` scripts

Add (keep existing migrate scripts unchanged):

```json
"db:snapshot": "node scripts/db/snapshot.mjs .env.development.local",
"db:snapshot:prod": "node scripts/db/snapshot.mjs .env.production.local",
"db:restore": "node scripts/db/restore.mjs .env.development.local"
```

Notes:

- No `db:restore:prod` script — ever.
- `db:migrate` remains the post-restore catch-up command.

### A5) Gitignore + docs

1. Add to `.gitignore`:

   ```gitignore
   app/db/snapshot.sql
   ```

2. Update `README.md` Neon section with a short “Resetting Neon dev” blurb:

   ```text
   npm run db:snapshot   # refresh baseline (dev by default)
   npm run db:restore    # wipe Neon dev + load snapshot
   npm run db:migrate    # catch up migrations
   ```

3. Optional one-liner in `RELEASE_CHECKLIST.md` under pre-release: restore+migrate is a valid local reset before validating migrations — not required for every release.

4. Keep this plan file as the detailed reference; link it from README if helpful.

### A6) Tests / validation

Add automated coverage for **safety guards** (required by repo learnings):

- Restore refuses a production env file path.
- URL resolution preference matches shared helper (can import `drizzle-migrate-shared.mjs` / lib helpers).
- Prefer unit-testing pure helpers; do not hit real Neon in CI.

Local manual checks (agent or human):

1. `npm run db:snapshot` → creates/updates `app/db/snapshot.sql` against **dev**.
2. `npm run db:restore` without typing `yes` → no wipe.
3. `npm run db:restore -f` (or confirm `yes`) → wipe + load on **dev** only.
4. `npm run db:migrate` → applies only newer migrations.
5. Confirm `git status` does **not** try to commit `app/db/snapshot.sql`.

Do **not** run restore against prod. Do **not** run `db:snapshot:prod` unless the human asked for it in hand-off.

### A7) Ship

1. Open a PR with scripts + gitignore + docs + tests.
2. PR body should remind: snapshot files stay local/private; restore is destructive to Neon **dev**.
3. Do not merge unless the user asks.

---

## Day-to-Day Usage (after scripts land)

### Refresh the baseline snapshot

When Neon **dev** (or a curated branch) is in a state you want to freeze:

```bash
npm run db:snapshot
# intentional prod-like dump only:
npm run db:snapshot:prod
```

### Reset a messy Neon dev DB

```bash
npm run db:restore
npm run db:migrate
```

### Fresh laptop / teammate setup

1. Put Neon URLs in `.env.development.local` / `.env.local`.
2. Obtain `app/db/snapshot.sql` privately (not via git if it has real data), **or** run `db:snapshot` from a shared Neon branch.
3. `npm run db:restore`
4. `npm run db:migrate`
5. `npm run dev`

### Alternate wipe: Neon branch reset

Instead of `DROP SCHEMA`:

1. Neon Console → **dev** branch → **Reset from parent** (or recreate the branch).
2. Then load snapshot if needed, or skip the dump if the parent already has the data you want.

Branch reset = “throw away my broken branch.” Snapshot files = “reload this exact baseline on any machine.”

---

## Risks / Guardrails

- **Never restore onto prod** — enforce in `db:restore`; no prod restore script.
- **PII / secrets** in dumps — treat `snapshot.sql` like a credential; gitignore it.
- **SSL** — Neon URLs should include `sslmode=require`. Fix the URL if clients fail; don’t casually disable verify.
- **Pooled vs direct** — prefer unpooled for dump/restore.
- **Migration watermark** — dumps include `__drizzle_migrations`. After restore, `db:migrate` only applies newer files. That is correct.

## Done When

- [ ] Human hand-off complete (`pg_dump`/`psql` + dev env direct URL)
- [ ] `scripts/db/snapshot.mjs` and `scripts/db/restore.mjs` exist with hostname-only logging
- [ ] `db:restore` refuses production env files and prompts unless `-f`
- [ ] `package.json` has `db:snapshot`, `db:snapshot:prod`, `db:restore`
- [ ] `app/db/snapshot.sql` is gitignored
- [ ] Safety-guard tests pass
- [ ] README documents the three-command reset loop
- [ ] Manual: snapshot → restore (dev) → migrate brings schema to HEAD without touching prod

## Related

- Production auto-migrate on deploy: `ai-assistance/PLAN_DB_MIGRATE_VERCEL.md`
