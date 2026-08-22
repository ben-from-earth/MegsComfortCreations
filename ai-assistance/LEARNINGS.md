# LEARNINGS

## Purpose

Capture durable corrections so agents avoid repeating mistakes and preserve long-term code quality.

## Entry Format

- Date:
- Area:
- Issue:
- Correction:
- Rule Going Forward:

## Entries

### 2026-03-04 - TypeScript typing approach

- Area: RBAC / TypeScript patterns
- Issue: Using `as` casts to silence type errors without solving the underlying type design issue.
- Correction: Prefer `satisfies`, explicit existing repo types, or new shared types/interfaces where needed so typing remains accurate and reusable.
- Rule Going Forward: Avoid `as` casting unless there is no safer alternative and the reason is documented inline. Default to modeling the types correctly so TypeScript keeps providing real safety instead of bypassed checks.

### 2026-03-08 - Descriptive naming over abbreviations

- Area: Variable and function naming
- Issue: Ambiguous abbreviations and short names reduce readability and make intent harder to understand during maintenance and review.
- Correction: Use descriptive names that communicate purpose and behavior for both values and functions, even when names are slightly longer.
- Rule Going Forward: Prefer explicit, intent-revealing names over abbreviations or shortcuts so each identifier clearly describes the represented value or action.

### 2026-03-15 - Regression coverage for shipped changes

- Area: Testing discipline
- Issue: Bug fixes and feature updates were not consistently accompanied by tests, allowing regressions to reappear.
- Correction: Modernized stale tests to align with active tRPC routers and added broad mocked regression coverage for current router behavior.
- Rule Going Forward: Any bug fix or feature change must include new or updated automated tests in the same task to lock in expected behavior and prevent regression.

### 2026-07-26 - Windows shell must not wrap Neon connection URLs

- Area: Local DB scripts (`pg_dump` / `psql`) on Windows
- Issue: Spawning Postgres clients with `shell: true` caused cmd.exe to split Neon URLs on `&` in query params (e.g. `channel_binding=require`), breaking snapshot/restore.
- Correction: Resolve the executable via `where.exe`/`which`, then `spawnSync` with `shell: false` so the full URL stays a single argv entry.
- Rule Going Forward: Never pass database connection URLs through a Windows shell. Prefer env-var handoff (as migrate scripts do) or argv with `shell: false`.

### 2026-07-26 - Explicit env files must override shell exports

- Area: `db:snapshot*` / `db:migrate*` env loading
- Issue: `dotenv` defaults to `override: false`, so a leftover `$env:DATABASE_URL` in the shell could make `db:snapshot:prod` dump the wrong Neon branch.
- Correction: Load script env files with `override: true` so the file passed on the CLI always wins.
- Rule Going Forward: Any script that takes an explicit env-file path must load it with `override: true` (or clear prior DB URL keys first).

### 2026-08-08 - Prefer slim business-logic tests

- Area: Testing discipline
- Issue: The suite accumulated long, mock-heavy UI and router tests that were slow, brittle, and hard to maintain.
- Correction: Cleared those suites and kept only fast pure-helper / safety-guard / slim auth-guard coverage until a deliberate testing approach is reintroduced.
- Rule Going Forward: Prefer slim, fast business-logic unit tests. Avoid broad mocked UI/component and heavily mocked router suites unless intentionally designed.

### 2026-08-20 - Do not extract helpers just to test them

- Area: Extraction / testing
- Issue: Tiny one-off field updates (exclusive image select, genre toggle, number coerce) were pulled into `collected-item-field-updates.ts` so they could have unit tests.
- Correction: Inline those operations at the call sites and delete the helper module and its tests.
- Rule Going Forward: Do not extract a helper solely to make it testable. Leave simple UI/form wiring inline. Extract only when the same logic is reused or the name itself is the abstraction.

### 2026-08-22 - No test-only helpers; skip unsolicited aria-labels

- Area: Extraction / testing / a11y
- Issue: Duplicate-book detection was pulled into `is-duplicate-book-error.ts` so it could have a unit test, and pager buttons were given aria-labels that were not requested.
- Correction: Inline the unique-constraint check in the `database.save` catch, delete the helper and its test, and drop the extra aria-labels.
- Rule Going Forward: Do not extract a helper just to test it. Tests stay slim and cover real business logic only — skip tests that exist only because something was extracted. Do not add aria-labels unless the user asks for them.
