# Codebase Review Issues

Date: 2026-03-08
Scope: Full repository review with bug risk/correctness priority, plus security and maintainability.

## Critical

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

### 8) Weak protection on internal admin seed endpoint

- **Path**
    - `app/api/internal/seed-admin/route.ts`
- **Issue**
    - Admin creation endpoint relies on one static header secret.
- **Fix suggestion**
    - Disable this endpoint in production or gate to explicit environment.
    - Add additional controls: short-lived signed token, rate limiting, source restriction, and generic error responses.

## Medium

### 12) Potential crash on missing env vars

- **Paths**
    - `app/db/client.ts`
    - `drizzle.config.ts`
- **Issue**
    - Non-null assertions (`process.env.DATABASE_URL!`) fail hard with poor diagnostics.
- **Fix suggestion**
    - Add startup env validation with explicit errors and required var checks.

## Cleanup and Maintainability Opportunities

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

## Suggested Fix Order (Practical)

1. Add server-side authz for tRPC mutations.
2. Fix `database.getPaginated` and `database.edit` type handling.
3. Fix enum/sort contract drift between UI and backend.
4. Add transaction safety for `database.save`.
5. Harden PNG fetch against SSRF.
6. Replace stale tests with tRPC-aligned integration coverage.
