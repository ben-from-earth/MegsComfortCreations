# Plan: Auto-Migrate Prod on Vercel Build

## Goal

Stop the “deploy succeeds, prod breaks until I manually migrate” loop.

- **Local/dev:** `npm run db:migrate` always targets Neon **dev** via `.env.development.local`.
- **Production:** pending Drizzle migrations run automatically during the Vercel **production** build (using Vercel env vars).
- **Manual prod migrate:** `npm run db:migrate:prod-emergency` only — laptop escape hatch.

## Target End State

```text
Local work:
  npm run db:generate          # if schema changed
  npm run db:migrate           # apply to Neon dev

Merge to main → Vercel production build:
  1) drizzle-kit migrate (prefer unpooled URL from Vercel env)
  2) next build
  3) deploy

If migrate fails → build fails → old prod keeps running (good)
Local npm run build → migrate no-ops (does not touch Neon)
```

## Already Done (do not redo)

- [x] `package.json` scripts:
  - `db:migrate` → `.env.development.local`
  - `db:migrate:prod-emergency` → `.env.production.local`
- [x] Docs partially renamed to those script names (still need happy-path rewrite below)

---

## Human Input Required

Do these **before** the agent merges / you ship the `build` script change (or in the same release window, before relying on auto-migrate).

The agent **cannot** log into Neon or Vercel for you. Complete every checkbox, then reply in chat with the filled “Hand-off back to agent” block.

### H1) Neon — copy connection strings (prod branch)

1. Open [Neon Console](https://console.neon.tech) → this project.
2. Select the **production** branch (usually `main` / production).
3. Open **Connection details**.
4. Copy **two** URLs for that same branch:
   - **Pooled** — hostname usually contains `-pooler` → this is app runtime `DATABASE_URL`.
   - **Direct / unpooled** — hostname has **no** `-pooler` → this is migrate `DATABASE_URL_UNPOOLED`.
5. Keep both private (do not paste full URLs into git or this plan file).

### H2) Vercel — set Production env vars

1. Open Vercel → this project → **Settings** → **Environment Variables**.
2. For **Production** only, ensure:

| Name | Value | Environments | Available at |
|------|--------|--------------|--------------|
| `DATABASE_URL` | Neon **pooled** prod URL | Production | Build + Runtime |
| `DATABASE_URL_UNPOOLED` | Neon **direct** prod URL | Production | Build (+ Runtime ok) |

3. If `DATABASE_URL` already exists for Production, only add/update `DATABASE_URL_UNPOOLED` unless the pooled URL is wrong.
4. Do **not** enable Preview for these prod URLs unless you intentionally want previews on prod DB (you do not).
5. Confirm neither var is restricted to Runtime-only (migrate runs at **build** time).
6. Optional: add `MIGRATE_ON_BUILD=true` for Production only if you want an explicit override flag in addition to `VERCEL_ENV === 'production'`. Not required if the agent gates on `VERCEL_ENV`.

### H3) Vercel — confirm Build Command

1. Vercel → **Settings** → **General** → **Build & Development Settings**.
2. Prefer **Build Command** empty / default so it uses `npm run build` from `package.json`.
3. If Build Command is overridden to plain `next build`, change it to `npm run build`.
4. Framework Preset: Next.js is fine.

### H4) Confirm policy (no action if you agree)

- Preview deployments must **not** migrate the production database.
- Agent will implement: migrate only when `VERCEL_ENV === 'production'` (or `MIGRATE_ON_BUILD=true`).
- Reply if you want a different policy.

### Hand-off back to agent

Paste this filled block when H1–H3 are done:

```text
Human setup complete:
- Neon prod branch name: <e.g. main>
- Vercel Production has DATABASE_URL (pooled): yes/no
- Vercel Production has DATABASE_URL_UNPOOLED (direct): yes/no
- Both available at Build time: yes/no
- Build Command uses npm run build (or default): yes/no
- Preview will NOT get prod DB URLs: yes/no
- Ready for agent to implement + open PR: yes/no
```

Do **not** paste secret connection strings here.

---

## Agent Implementation Steps

Execute in order. Stop and ask the human if a step conflicts with repo state.

### A0) Preconditions

1. Read `ai-assistance/LEARNINGS.md` and this plan.
2. Confirm human hand-off block is complete (or proceed with code-only PR and note that Vercel/Neon env must be set before merge).
3. Confirm existing helpers:
   - `drizzle-migrate-with-env.mjs` (local file-based migrate)
   - `drizzle.config.ts` already prefers `DRIZZLE_DATABASE_URL` then `DATABASE_URL`

### A1) Create `drizzle-migrate-ci.mjs`

New file at repo root (same style as `drizzle-migrate-with-env.mjs`).

Behavior:

1. **Gate:** if `process.env.VERCEL_ENV !== 'production'` **and** `process.env.MIGRATE_ON_BUILD !== 'true'`:
   - Log a one-line skip reason (e.g. `Skipping Drizzle migrate (not production build)`).
   - `process.exit(0)` — do not fail local/preview builds.
2. **Resolve URL** (first present wins):
   1. `DRIZZLE_DATABASE_URL`
   2. `DATABASE_URL_UNPOOLED`
   3. `DATABASE_URL`
3. If none set after the gate passes: log clear error and `process.exit(1)`.
4. Set `process.env.DRIZZLE_DATABASE_URL` to the resolved URL.
5. Log **hostname only** via `new URL(...).hostname` — never log the full URL/password.
6. Prefer logging whether the source was unpooled vs fallback (without printing secrets), e.g. `Using DATABASE_URL_UNPOOLED`.
7. Spawn `npx drizzle-kit migrate` the same cross-platform way as `drizzle-migrate-with-env.mjs` (`spawnSync`, win32 shell handling).
8. Exit with the migrate process status; non-zero on failure.

Acceptance:

- Local `node drizzle-migrate-ci.mjs` with no special env → exits 0 and skips.
- With `MIGRATE_ON_BUILD=true` and a missing URL → exits non-zero.
- Never prints connection string credentials.

### A2) Wire `package.json` `build`

Set:

```json
"build": "node drizzle-migrate-ci.mjs && next build"
```

Keep:

```json
"db:migrate": "node drizzle-migrate-with-env.mjs .env.development.local",
"db:migrate:prod-emergency": "node drizzle-migrate-with-env.mjs .env.production.local"
```

Do **not** add a `db:migrate:ci` script.

Acceptance:

- `npm run build` locally still completes (migrate no-ops, then `next build`).
- No accidental migrate against Neon during a normal local build.

### A3) (Optional but recommended) Align local emergency migrate with unpooled preference

Only if cheap and consistent: in `drizzle-migrate-with-env.mjs`, resolve migrate URL as:

1. `DRIZZLE_DATABASE_URL` (if already set)
2. `DATABASE_URL_UNPOOLED` from the loaded env file
3. `DATABASE_URL` from the loaded env file

Then set `DRIZZLE_DATABASE_URL` from that. Keep requiring at least one of unpooled/pooled.

Do **not** change `app/db/client.ts` — runtime stays on pooled `DATABASE_URL`.

### A4) Update docs

#### `ai-assistance/RELEASE_CHECKLIST.md`

Rewrite happy path to:

1. Local: generate (if needed) → `npm run db:migrate` → validate.
2. Merge to `main`.
3. Vercel production build runs migrate automatically.
4. Smoke-test prod.

Move laptop prod migrate to an **Emergency only** section:

```bash
npm run db:migrate:prod-emergency
```

Remove / stop recommending dead `db:check:*` commands (script file was deleted). Do not restore them in this task unless the user asks.

#### `README.md` — Neon Branch Safety section

Update to say:

- Production migrations run on Vercel production build.
- Local: `npm run db:migrate` (dev).
- Emergency only: `npm run db:migrate:prod-emergency`.
- Remove stale `db:check:*` command list.

#### `.env.example`

Document:

```env
# App runtime (pooled Neon URL is fine)
DATABASE_URL=postgresql://user:password@host-pooler/db?sslmode=require

# Preferred for migrations (direct / non-pooler Neon URL)
# Used by drizzle-migrate-ci.mjs on Vercel; optional locally
DATABASE_URL_UNPOOLED=postgresql://user:password@host/db?sslmode=require

# Local migration env files (not committed):
# - .env.development.local  → npm run db:migrate
# - .env.production.local   → npm run db:migrate:prod-emergency
```

Remove outdated `db:migrate:dev` / `db:migrate:prod` / `db:check:*` comments.

### A5) Out of scope unless user asks

- Fixing Drizzle `app/db/migrations/meta/` snapshot lag (journal to `0007`, meta only through `0003`) — note in PR if noticed; do not block this change.
- Neon + Vercel preview DB branching.
- Restoring `db:check:*` tooling.

### A6) Validation (agent)

Run what you can locally:

1. `node drizzle-migrate-ci.mjs` → skip, exit 0.
2. `MIGRATE_ON_BUILD=true node drizzle-migrate-ci.mjs` without URL → exit non-zero (or with a throwaway invalid URL — do **not** point at real prod).
3. Targeted tests / lint / typecheck for touched areas if applicable.
4. Prefer not running a full `next build` unless needed; if run, confirm migrate skipped first.

Do **not** run `db:migrate:prod-emergency` against real prod as part of implementation.

### A7) Ship

1. Open a PR with the code + doc changes.
2. In the PR body, remind the human:
   - Confirm Vercel env hand-off is done before merging.
   - After merge, check production deploy logs for hostname-only migrate line, then smoke-test.
3. Do not merge unless the user asks.

---

## Post-Merge Verification (human + agent)

After merge to `main`:

1. Open the Vercel **production** deployment log.
2. Confirm a line like: `Running Drizzle migrate against host: ep-....neon.tech` (direct host preferred; should not be the only path if unpooled is set).
3. Confirm migrate completes **before** `next build`.
4. Smoke-test prod.
5. Optional: Neon SQL editor on prod → `__drizzle_migrations` includes latest journal entries.

If auto-migrate fails:

```bash
# Laptop; requires .env.production.local with prod URLs
npm run db:migrate:prod-emergency
```

Then fix the CI migrate path and redeploy.

---

## Local Day-to-Day Workflow (after change)

1. Edit `app/db/schema.ts`.
2. `npm run db:generate` — commit SQL under `app/db/migrations/`.
3. `npm run db:migrate` — apply to Neon **dev**.
4. Test locally.
5. Merge to `main`.
6. Vercel production build migrates Neon **prod**.
7. Do **not** use `db:migrate:prod-emergency` on a normal release.

---

## Risks / Guardrails

- Destructive migrations run automatically on deploy — review SQL in PR.
- Failed migrate fails deploy — correct; previous prod stays up.
- Wrong Vercel URL — verify hostname in first production migrate log.
- Local `npm run build` must no-op migrate unless `MIGRATE_ON_BUILD=true`.
- Preview builds must not migrate production.

## Done When

- [ ] Human hand-off block completed (Vercel/Neon env + build command)
- [ ] `drizzle-migrate-ci.mjs` exists with production/MIGRATE_ON_BUILD gate + URL resolution order
- [ ] `package.json` `build` runs migrate-ci then `next build`
- [ ] Local build skip behavior verified
- [ ] `RELEASE_CHECKLIST.md`, `README.md`, `.env.example` updated (no manual-prod happy path; no dead `db:check:*`)
- [ ] PR opened; first production deploy log shows migrate against expected host
- [ ] Preview builds cannot migrate production
