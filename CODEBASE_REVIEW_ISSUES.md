# Codebase Review Issues

Date: 2026-03-08
Scope: Full repository review with bug risk/correctness priority, plus security and maintainability.

## Critical

### 1) Unauthenticated tRPC mutations and data access (COMPLETED)
- **Paths**
  - `lib/trpc/context.ts`
  - `lib/trpc/trpc.ts`
  - `lib/trpc/routers/_app.ts`
  - `lib/trpc/routers/database/_.ts`
  - `lib/trpc/routers/collect/_.ts`
  - `lib/trpc/routers/genres/_.ts`
  - `lib/trpc/routers/png.ts`
- **Issue**
  - Most procedures are exposed via `publicProcedure` with no server-side session/role checks.
- **Fix suggestion**
  - Add auth/session to tRPC context.
  - Implement `protectedProcedure` and `adminProcedure` middleware.
  - Restrict write/destructive procedures (save/edit/delete/link/unlink/png/create/collect) to authenticated roles.

### 2) `getPaginated` ignores media type and always queries books
- **Path**
  - `lib/trpc/routers/database/_.ts`
- **Issue**
  - API accepts `type` (`book|movie|videoGame|album`) but query logic is book-specific.
- **Fix suggestion**
  - Branch by `type` with dedicated query builders per table.
  - Align selected columns and sort options to each media type.

### 3) SSRF risk in PNG image fetching
- **Paths**
  - `lib/helpers/outputPNG.ts`
  - `lib/trpc/routers/png.ts`
- **Issue**
  - Server fetches user-provided URLs (`axios.get`) without allowlist/private-network blocking.
- **Fix suggestion**
  - Add strict domain allowlist and block loopback/link-local/private ranges.
  - Revalidate redirect targets and cap redirects.
  - Prefer fetching only pre-vetted image URLs stored by your backend.

## High

### 4) Pagination total count ignores active filters (COMPLETED)
- **Path**
  - `lib/trpc/routers/database/_.ts`
- **Issue**
  - `total` is computed from all books, not the filtered query.
- **Fix suggestion**
  - Build a mirrored count query using the same `where` filters and joins as the page query.

### 5) Genre "None" mismatch between UI and backend (COMPLETED)
- **Paths**
  - `app/showdatabase/PaginationInputs.tsx`
  - `lib/trpc/routers/database/_.ts`
- **Issue**
  - UI sends `'none'`, backend expects `'None'`.
- **Fix suggestion**
  - Centralize shared filter constants/enums in one module and consume from both UI and router input schema.

### 6) `database.edit` ignores `type` and updates books only
- **Path**
  - `lib/trpc/routers/database/_.ts`
- **Issue**
  - Non-book edits are accepted by API contract but not actually handled correctly.
- **Fix suggestion**
  - Split edit handling by media type with per-type validation schema and table mapping.

### 7) Partial writes possible during save
- **Path**
  - `lib/trpc/routers/database/_.ts`
- **Issue**
  - Book insert + genre link inserts are not transactional; failures can leave partial state.
- **Fix suggestion**
  - Wrap insert and genre-link loop in a DB transaction.
  - Validate `genreRow` presence before insert; fail fast with clear error result.

### 8) Weak protection on internal admin seed endpoint
- **Path**
  - `app/api/internal/seed-admin/route.ts`
- **Issue**
  - Admin creation endpoint relies on one static header secret.
- **Fix suggestion**
  - Disable this endpoint in production or gate to explicit environment.
  - Add additional controls: short-lived signed token, rate limiting, source restriction, and generic error responses.

## Medium

### 9) Pagination page-change fetch race
- **Paths**
  - `app/shared/PageSelector.tsx`
  - `app/showdatabase/page.tsx`
- **Issue**
  - `setPage(...)` followed by immediate fetch can use stale page state.
- **Fix suggestion**
  - Remove immediate manual fetch from button handlers.
  - Let query react to state change or pass explicit next page to refetch logic.

### 10) Sort option drift causes unexpected ordering
- **Paths**
  - `app/showdatabase/PaginationInputs.tsx`
  - `app/showdatabase/page.tsx`
  - `lib/trpc/routers/database/_.ts`
- **Issue**
  - UI and backend sort contracts do not fully match (`pageCount` and fallback behavior mismatch).
- **Fix suggestion**
  - Use one shared sort contract source for both client and server.
  - Reject unsupported sorts explicitly instead of silent fallback.

### 11) Query usage counter can lose updates
- **Path**
  - `lib/trpc/routers/collect/_.ts`
- **Issue**
  - Read-modify-write for daily count is non-atomic and may no-op if row missing.
- **Fix suggestion**
  - Use atomic increment/upsert in a transaction (or single SQL upsert statement).

### 12) Potential crash on missing env vars
- **Paths**
  - `app/db/client.ts`
  - `drizzle.config.ts`
- **Issue**
  - Non-null assertions (`process.env.DATABASE_URL!`) fail hard with poor diagnostics.
- **Fix suggestion**
  - Add startup env validation with explicit errors and required var checks.

## Cleanup and Maintainability Opportunities

### 13) Dead/commented component file
- **Path**
  - `app/mediacollector/MediaCheckboxes.tsx`
- **Issue**
  - Entire file is commented-out dead code.
- **Fix suggestion**
  - Delete file and stale commented imports/usages.

### 14) Likely unused component
- **Path**
  - `app/mediacollector/MediaInputs.tsx`
- **Issue**
  - Appears not wired into the active page flow.
- **Fix suggestion**
  - Remove if obsolete, or integrate and cover with tests.

### 15) Brittle direct node_modules import
- **Path**
  - `lib/trpc/routers/collect/actions/get-media-covers.ts`
- **Issue**
  - `import axios from 'node_modules/axios'`.
- **Fix suggestion**
  - Replace with `import axios from 'axios'`.

### 16) Stale ESLint override path
- **Path**
  - `eslint.config.mjs`
- **Issue**
  - Override references `src/app/...` while project uses `app/...`.
- **Fix suggestion**
  - Update override path or remove obsolete override.

### 17) Legacy helper script with hardcoded password string
- **Path**
  - `lib/helpers/createPasswordHash.ts`
- **Issue**
  - Hardcoded password-like value and hash logging pattern.
- **Fix suggestion**
  - Remove from repo or convert to a safe local CLI utility with env input and no secret logging.

## Test and Documentation Gaps

### 18) Tests appear stale vs current API architecture
- **Paths**
  - `__tests__/medias.test.ts`
  - `__tests__/genres.test.ts`
  - `app/api/trpc/[trpc]/route.ts`
- **Issue**
  - Tests target legacy REST-style route imports and outdated enum naming (`video_game` vs `videoGame`).
- **Fix suggestion**
  - Replace with tRPC integration tests for active routers and procedures.
  - Keep tests aligned to shared type contracts.

### 19) Type-check excludes tests
- **Path**
  - `tsconfig.json`
- **Issue**
  - `__tests__` excluded from TypeScript checks.
- **Fix suggestion**
  - Include tests in a dedicated test tsconfig (CI) or include them in main typecheck path.

### 20) `.env.local` not ignored by git
- **Path**
  - `.gitignore`
- **Issue**
  - `.env` ignored, but `.env.local` not explicitly ignored.
- **Fix suggestion**
  - Add `.env.local` and optionally `.env.*` patterns to `.gitignore`.
  - Rotate secrets if any sensitive values were committed/shared previously.

## Suggested Fix Order (Practical)

1. Add server-side authz for tRPC mutations.
2. Fix `database.getPaginated` and `database.edit` type handling.
3. Fix enum/sort contract drift between UI and backend.
4. Add transaction safety for `database.save`.
5. Harden PNG fetch against SSRF.
6. Replace stale tests with tRPC-aligned integration coverage.
