# Media Types Refactor Plan

## Goal

Simplify media handling to two data models:

- `books` for book-specific data (`author`, `pageCount`, `pubYear`, etc.)
- `other_media` for non-book types (`movie`, `videoGame`, `album`) with a `mediaType` discriminator

This reduces branching and keeps the codebase aligned with the real data differences.

## Target Data Model

- Keep existing `books` table (book-specific fields stay here).
- Replace `movies`, `video_games`, `albums` with one `other_media` table:
    - `id`
    - `mediaType` (`movie | videoGame | album`)
    - `title`
    - `spineColor`
    - `imageUrls`
- Add uniqueness on (`mediaType`, `title`) to prevent duplicate titles within the same non-book media type while still allowing same title across different media types.
- Add DB-level guardrails for `other_media`:
    - Constrain `mediaType` to allowed values (`movie | videoGame | album`) via enum/check constraint.
    - Add indexes for common query patterns (`mediaType`, title search/sort).
    - Keep `imageUrls` not-null behavior consistent with existing non-book tables.

## Refactor Scope

1. Database schema and migration

- Add `other_media` table and migration.
- Backfill/migrate rows from `movies`, `video_games`, and `albums`.
- Use a two-phase migration file:
    - Phase 1: create and backfill `other_media` from `movie`, `video_game`, and `album` tables.
    - Phase 2: drop old per-type non-book tables.
- Add migration validation checks before table removal:
    - Verify row counts between source tables and `other_media`.
    - Run spot checks for representative rows and key fields (`title`, `spineColor`, `imageUrls`, `mediaType`).
    - Define fallback handling for dirty/invalid legacy rows encountered during backfill.

2. Shared types and validation contracts

- Update shared media unions in `lib/interfaces/globalInterfaces.ts`.
- Update collector schema in `app/mediacollector/collector-form/collectorFormSchema.ts`.
- Define two block-info contracts:
    - Book block info (book fields)
    - Other media block info (title + spineColor + genres behavior as defined)

3. tRPC router cleanup

- Refactor `lib/trpc/routers/database/_.ts`:
    - `getPaginated`: query `books` or `other_media` based on `type`
    - `save`: insert into `books` or `other_media` with proper validation
    - `edit`: update correct table by `type`
    - `deleteByTitle`: delete from correct table by `type`
- Refactor `lib/trpc/routers/database/actions/search-by-title.ts` to use the two-table mapping.
- Refactor `lib/trpc/routers/collect/_.ts` to stop forcing book-shaped block data for all media.
- Wrap `save` writes in transactions where multi-step writes occur (book insert + genre links) to avoid partial writes.

4. UI alignment

- Update `app/showdatabase/PaginationInputs.tsx` sort/filter contracts to match backend capabilities.
- Update `app/showdatabase/EditDatabaseBlock.tsx` and related components so fields are type-correct:
    - Books: title, author, pubYear, pageCount, spineColor
    - Other media: title, spineColor
- Keep book-only genre link behavior isolated to books unless expanded later.
- Update collector form/input flow so non-book collection is explicitly supported end-to-end (schemas, defaults, and UI input wiring), not just backend type acceptance.

5. Dead code removal and consistency cleanup

- Remove commented legacy media files/branches (including `app/mediacollector/MediaCheckboxes.tsx` if not needed).
- Normalize naming (`videoGame` everywhere; remove legacy `video_game` usage in active paths).
- Normalize codebase to use the following as the source of truth for media types
    - [
      {mediaType: 'book', label: 'Book', plural: 'Books'},
      {mediaType: 'movie', label: 'Movie', plural: 'Movies'},
      {mediaType: 'videoGame', label: 'Video Game', plural: 'Video Games'},
      {mediaType: 'album', label: 'Album', plural: 'Albums'},
      ]

6. Tests and docs

- Update tests to current tRPC architecture and two-table model.
- Add/adjust coverage for `collect`, `save`, `edit`, and `getPaginated` across all media types.
- Add migration-focused tests/checks for backfill correctness and rollback safety.
- Update docs that still imply the old separate non-book table architecture.
- Update Swagger/OpenAPI docs to align naming/contracts (`videoGame` vs legacy `video_game`) and new table behavior.
- Update operational scripts/docs that assume legacy non-book tables (for example DB target-check workflows) to include `other_media`.

## Execution Order

1. Create schema migration files (`other_media` + backfill). I will manually migrate the databases after verification.
2. Update backend routers/types (including `deleteByTitle`) to read/write the new structure.
3. Update collector and database UI contracts/components.
4. Run validation checks (data integrity, lint, typecheck, tests) and fix regressions.
5. Update docs/scripts (Swagger + operational DB workflows) to new contracts.
6. Remove dead/commented code and legacy naming remnants or translate them back into the app if their original purpose is re-needed.

## Success Criteria

- All media types function end-to-end with only two models (`books`, `other_media`).
- No endpoint accepts a type it cannot correctly process.
- UI and backend share the same sort/filter/type contract.
- Commented legacy branches are removed from active code.
- Tests validate both book and non-book paths.
