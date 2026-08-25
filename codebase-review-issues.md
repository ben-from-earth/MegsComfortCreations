# Codebase Review Issues

Date: 2026-07-26  
Updated from: 2026-03-08 review + full codebase audit  
Scope: Remaining open/partial issues, plus typing, duplication, structure, dead code, and bloat.

---

## Resolved (removed from active list)

These were open in the March review and are now fixed:

- DB env validation (`app/db/client.ts`, `drizzle.config.ts`)
- Stale REST tests migrated to tRPC (`__tests__/medias.test.ts`, `__tests__/genres.test.ts`, `__tests__/trpc/*`)
- Server-side tRPC mutation authz (`lib/trpc/trpc.ts` `adminProcedure` + `__tests__/trpc/auth-guards.test.ts`)
- Enum/sort contract drift (`lib/constants/database-sort-options.ts`, `videoGame` naming)
- Internal admin seed endpoint removed (`app/api/internal/seed-admin/route.ts`) — bootstrap no longer needed; admins already exist in all environments
- Concurrent-delete edge in `database.edit` — empty update `returning` now yields `Media Not Found` instead of asserting on the row
- `database.save` DB writes use transactions + compensating parent delete on S3 failure; S3 orphans on rare post-upload DB failure remain accepted (S3 cannot join a Postgres txn). User-facing image failure copy no longer exposes raw S3/URL details.
- Swagger / OpenAPI stack removed (`/docs`, `swagger-ui-react`, `swagger-jsdoc`, related types, Redux peers kept only for Swagger, stale `tsconfig` doc path aliases, README `/docs` section)
- Low-value / stale test cleanup: removed duplicate smoke suites and placeholder `profile` / `health` router coverage; `auth-guards` now asserts against a real `adminProcedure` (`genres.getAll`); unused `health` router deleted

---

## Security

### SSRF risk in PNG image fetching (partial)

- **Paths**
    - `lib/helpers/output-png.ts`
    - `lib/trpc/routers/png.ts`
- **Status**
    - Timeouts, payload caps, and `maxRedirects: 3` exist.
    - Still missing hostname allowlist, private/loopback/link-local blocking, and redirect-target revalidation.
- **Fix suggestion**
    - Allowlist trusted image hosts (or only fetch backend-stored URLs).
    - Resolve DNS and block private ranges before fetch; revalidate each redirect hop.

---

## Typing & Shared Contracts

### Drifted `MediaImage` shapes across layers

- **Paths**
    - `lib/interfaces/global-interfaces.ts` (`MediaImageItem`)
    - `app/components/shared/media-image-strip.tsx` (`MediaImageStripItem`)
    - `app/mediacollector/collector-form/collector-form-schema.ts` (`imageSelectionSchema`)
    - `lib/media-storage/media-image-records.ts`
- **Issue**
    - Same concept (url / selected / isDefault / spineColor) with conflicting optionality and field rules.
- **Fix suggestion**
    - One base domain type (or Zod schema) with derived UI / persist / editable variants.

### Collector form block-info overlap

- **Paths**
    - `app/mediacollector/collector-form/media-item-form-schema.ts` (`mediaItemBlockInfoSchema`)
    - `app/mediacollector/collector-form/collector-form-schema.ts`
- **Issue**
    - Parallel models with different nullability and fields. The unused `BlockInfo` interface was removed.
- **Fix suggestion**
    - Treat Zod schemas as canonical; derive display/persist types via `z.infer` / `Pick` / `.pick()`.

### Duplicated collection-list + collected-block types

- **Paths**
    - `app/mediacollector/collector-form/collector-form-schema.ts`
    - `lib/trpc/routers/collect/_.ts`
- **Issue**
    - `{ title, author? }` arrays repeated for each media type in form schema and router input.
    - `CollectedBlockInformation` re-inferred in the router instead of imported.
- **Fix suggestion**
    - Export `collectionListSchema` / title-item schema once; import in form + router.
    - Import `CollectedBlockInformation` from the form schema module.

### Duplicated PNG export image contract

- **Paths**
    - `app/mediacollector/png-export-images.ts` (`PNGExportImage`)
    - `lib/helpers/output-png.ts` (`ImageData`)
- **Issue**
    - Same render-input shape defined twice.
- **Fix suggestion**
    - Shared server-safe type/schema; infer router input from it.

### Forced / redundant casts

- **Paths**
    - `app/showdatabase/page.tsx` — `paginatedQuery.data as SuccessfulPaginationResponse`
    - `app/showdatabase/pagination-inputs.tsx` — DOM string casts to `MediaType` / sort / genre / limit / direction
    - `lib/trpc/routers/collect/_.ts` — `key as MediaType` after `Object.entries`
    - `lib/trpc/routers/genres/_.ts` — DB `string[]` cast to `GenreEnum[]`; redundant `input.sort as SortKey` after Zod
    - `lib/media-storage/media-image-records.ts`, `lib/media-storage/image-path-utils.ts` — unknown object casts for url/src
    - `lib/trpc/routers/collect/actions/get-open-library-data.ts`, `lib/trpc/routers/online.ts` — Axios responses annotated, not Zod-parsed
- **Fix suggestion**
    - Prefer typed option arrays, `MEDIA_TYPES` iteration, Zod parse for external payloads, and tRPC-inferred response types over `as` casts.

---

## Code Duplication

### Collect router normalization branches

- **Path**
    - `lib/trpc/routers/collect/_.ts`
- **Issue**
    - Existing/non-existing × book/non-book paths repeat image normalization, defaults, block IDs, spine color, and block assembly.
- **Fix suggestion**
    - Extract `normalizeCollectedImages`, `createCollectedBlock`, and a small book-details branch.

### Book vs other-media image persistence

- **Path**
    - `lib/media-storage/media-image-records.ts`
- **Issue**
    - `replaceBookImageRecords` / `replaceOtherMediaImageRecords` (and related loaders/adapters) differ mainly by foreign key.
- **Fix suggestion**
    - Parameterize parent-id column/projection in shared helpers; keep thin named wrappers.

### Image-edit UI copied between collector and database editor

- **Paths**
    - `app/mediacollector/cbb-images.tsx`
    - `app/showdatabase/media-item-add-edit.tsx`
- **Issue**
    - Selection, EyeDropper color updates, base64 conversion, upload mutation payloads, and appended image state are near duplicates.
- **Fix suggestion**
    - Shared `fileToBase64`, selection updater, and upload/color hook; keep only context-specific commits in each component.

### Online cover fetch vs collect actions

- **Paths**
    - `lib/trpc/routers/online.ts`
    - `lib/trpc/routers/collect/actions/get-open-library-data.ts`
    - `lib/trpc/routers/collect/actions/get-media-covers.ts`
- **Issue**
    - Open Library / Google cover retrieval logic is duplicated.
- **Fix suggestion**
    - Route `online` through the shared collect actions.

### PNG layout algorithms

- **Path**
    - `lib/helpers/output-png.ts`
- **Issue**
    - Single-page and paged slot-layout logic are largely duplicated.
- **Fix suggestion**
    - Parameterize start index / capacity into one layout helper.

---

## File Structure & Conventions

### Split UI / shared / constants locations

- **Paths**
    - `app/shared/`
    - `app/components/ui/`
    - `app/lib/` (e.g. `app/lib/enums/genre-enums.ts`, `app/lib/utils/classnames.tsx`)
    - `lib/` (constants, helpers, media-storage, trpc)
- **Issue**
    - No single convention for shared UI vs domain libs; genre enums live under `app/lib` while other constants are under `lib/constants`.
- **Fix suggestion**
    - Prefer `app/components` (feature + shared UI) and root `lib` for non-UI.
    - Move `genre-enums` into `lib/constants/`.

### Stale agent / README structure docs

- **Paths**
    - `AGENTS.md` (references missing `CLAUDE.md`)
    - `README.md`
- **Issue**
    - README still describes REST folders, Redux, `.env`, and `jest.config.ts`; current stack is tRPC, no Redux, `.env.example`, `jest.config.mjs`.
- **Fix suggestion**
    - Add `CLAUDE.md` or drop the reference; rewrite README structure to match the repo.

## Dead Code & Unused Exports

Previously listed leftovers (`vanillaClient.ts`, `createPasswordHash.ts`, JSON create schemas, `Updates.txt`, unused sort-option exports, unused error types, seed SQL) have been removed.

---

## Dependency & Documentation Bloat

### Unused or likely-unused packages

- **Path**
    - `package.json`
- **Remaining candidate** (no project source imports found; keep as a Drizzle/Neon peer)
    - `@neondatabase/serverless`
- **Removed in this pass**
    - `jsonschema`, `bcrypt`, `@types/bcrypt`, `jsonwebtoken`, `@types/jsonwebtoken`, `msw`, `@mswjs/interceptors`, `@testing-library/react`, `@testing-library/jest-dom`, `jest-environment-jsdom`, `whatwg-url`

### Stale / low-value docs

- **Paths**
    - `docs/PROJECT_IDEAS.md`, `docs/PROJECT_PROPOSAL.md` — capstone-era material (already gone)
    - `docs/CONTROL_FLOW.md`, `docs/DATA_MODEL.md` — already gone
- **Fix suggestion**
    - No remaining stale docs in those paths.

---

## Tests

Need big time test rework

### Coverage gaps & CI friction

- **Paths**
    - `__tests__/trpc/png.router.test.ts` — mocks `outputAuto`; Sharp/ZIP path untested
    - `__tests__/trpc/genres.router.test.ts` — still uses brittle nested Drizzle fluent mocks (`database.router` now uses `__tests__/helpers/mockDrizzle.ts`)
    - `tsconfig.json` — excludes `__tests__`
    - `package.json` `test:ci` — `NODE_ENV=test` is POSIX-only (fails on Windows)
- **Fix suggestion**
    - Migrate remaining router tests to `createMockDb` / `createChainableQueryMock`.
    - Add focused integration coverage for critical DB/PNG paths.
    - Wire `__tests__/tsconfig.json` (or equivalent) into CI typecheck.
    - Use `cross-env` or set `NODE_ENV` inside Jest config for `test:ci`.

---

## Suggested Fix Order

1. Harden PNG fetch against SSRF.
2. Unify `MediaImage` / collector block-info / `collectionList` schemas as single sources of truth; remove redundant casts.
3. Deduplicate collect normalization, image persistence helpers, and cbb-images / media-item-add-edit image-edit logic.
4. Refresh README / AGENTS structure docs for the current tRPC stack.
5. Fix Windows `test:ci`; add test typecheck to CI.
