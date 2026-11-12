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

### 2026-08-25 - Kebab-case filenames

- Area: File naming
- Issue: The repo mixed PascalCase components, camelCase helpers, and kebab-case newer modules, which made imports inconsistent.
- Correction: Renamed source, helper, test-helper, and remaining docs to kebab-case. Left Next.js reserved files, tool configs, Drizzle migrations, route directories, and public image URLs unchanged.
- Rule Going Forward: New files use kebab-case (`media-checkboxes.tsx`) unless the name is required by Next.js, a tool, or a public URL. Exported symbols stay PascalCase/camelCase; only the filename changes.

### 2026-10-21 - PNG order is not collector UX order

- Area: Media collector PNG export
- Issue: A proposal treated “PNG should follow the customer list” as a reason to also stop grouping database items at the bottom of the collector grid.
- Correction: Collector UX stays database-last. Only the PNG is rebuilt in pasted-list order. Those are two lists.
- Rule Going Forward: Do not change collector display order when fixing PNG order. Keep `collectedData` as the UX list. Drive PNG order from a separate encounter-order `blockID` list.

### 2026-10-30 - User handles browser testing

- Area: Validation workflow
- Issue: The agent opened the app in the browser to verify UI changes; the user does that review themselves.
- Correction: Stop at lint, typecheck, and tests, then hand the browser check to the user.
- Rule Going Forward: Do not run browser testing unless the user asks. Validate with lint, typecheck, and unit tests, and list what the user should check by hand.

### 2026-10-30 - Trace new fields through to form state

- Area: Collector data flow
- Issue: A `palette` field added to collected images was dropped on the client because `page.tsx` rebuilds images with `toFormImages`, which copies a fixed list of fields.
- Correction: Carry `palette` through `toFormImages` and add a test for it.
- Rule Going Forward: When adding a field to server responses, follow it through every client mapping into the form state before calling it done.

### 2026-10-31 - Do not slow collector search for a progress bar

- Area: Collector search progress
- Issue: Progress was implemented by searching one title per HTTP request so the modal could update, which added a network round trip between each title.
- Correction: Keep one server-side collect loop and stream progress events from that existing work.
- Rule Going Forward: Do not split a sequential server job into extra client requests just to drive a progress bar. Report progress from the work that is already happening.

### 2026-11-02 - Search results can list deleted files

- Area: Codebase search / scripted edits
- Issue: Search returned stale forward-slash paths for the removed `app/showdatabase/` page, so the agent reported it as still existing and ran a bulk PowerShell edit against it. The failed read left the previous file's text in the variable, and the write created a bogus `are-you-sure.tsx`.
- Correction: Deleted the stray file and confirmed paths with `Test-Path` / `git status`.
- Rule Going Forward: Confirm a file exists on disk before calling it live or editing it. Prefer the edit tool over read/write shell loops; if a script is needed, skip the write when the read fails.

### 2026-11-02 - Pick per-route visuals before paint

- Area: Nav paint-smear brush
- Issue: The brush was picked in `useEffect` on `pathname`, so the newly active tab painted one frame with the previous brush, and lazy, optimizer-served PNGs popped in late.
- Correction: Pick in `useLayoutEffect`, serve the brushes `unoptimized` with eager loading, and `preload` all of them.
- Rule Going Forward: When client-only state must change in step with a route change, update it in `useLayoutEffect` (not `useEffect`) and preload any images it swaps between.

### 2026-11-02 - Key nav visuals on the active tab, not the URL

- Area: Nav paint-smear brush
- Issue: The brush re-picked on every `pathname` change, so moving between shop pages repainted the Shop tab even though it stayed active.
- Correction: Derive the active tab (`'shop'` for any shop page, otherwise the matching link href) and key the pick on that.
- Rule Going Forward: Nav state tied to "the current tab" depends on the active tab, not the raw pathname; several routes can share one tab.

### 2026-11-11 - Tests follow the functions the app calls

- Area: Movie title search
- Issue: `movieTitlesMatch` existed only so a test could assert year matching. Collector search makes that decision in SQL and never called the helper. A second test restated the library filter by calling two helpers itself.
- Correction: Deleted the helper and that test. The remaining tests call `splitMovieTitle`, `toStoredMovieTitle`, `toDisplayMovieTitle`, and `movieTitleWithAnyYearPattern`, which save and search use.
- Rule Going Forward: Do not add a function whose only caller is a test. Test the functions the application calls. If a rule lives in a query, do not invent a twin for the test file.

### 2026-11-12 - Edited migrations do not reach an existing database

- Area: Drizzle migrations / local DB
- Issue: The library proposal said a changed migration hash would fail the next migrate. `drizzle-kit migrate` only compares timestamps, so it reported success and skipped the edited `0000`/`0004`/`0008`, leaving `year` and `console` missing. The local branch also held the old uniques as unique indexes, not constraints, so `DROP CONSTRAINT` failed.
- Correction: Snapshotted, then applied the column, title-split, and constraint changes in one transaction, dropping the old uniques with `DROP INDEX`.
- Rule Going Forward: After editing an applied migration, check the live columns. `db:migrate` succeeding proves nothing. Inspect `pg_indexes` and `pg_constraint` before dropping a unique, and snapshot before any hand-applied change.
