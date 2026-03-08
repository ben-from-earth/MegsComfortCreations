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
- Add uniqueness on (`mediaType`, `title`) to avoid cross-type title collisions.

## Refactor Scope

1. Database schema and migration
- Add `other_media` table and migration.
- Backfill/migrate rows from `movies`, `video_games`, and `albums`.
- Remove old per-type non-book tables after successful migration.

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
- Refactor `lib/trpc/routers/database/actions/search-by-title.ts` to use the two-table mapping.
- Refactor `lib/trpc/routers/collect/_.ts` to stop forcing book-shaped block data for all media.

4. UI alignment
- Update `app/showdatabase/PaginationInputs.tsx` sort/filter contracts to match backend capabilities.
- Update `app/showdatabase/EditDatabaseBlock.tsx` and related components so fields are type-correct:
  - Books: title, author, pubYear, pageCount, spineColor
  - Other media: title, spineColor
- Keep book-only genre link behavior isolated to books unless expanded later.

5. Dead code removal and consistency cleanup
- Remove commented legacy media files/branches (including `app/mediacollector/MediaCheckboxes.tsx` if not needed).
- Normalize naming (`videoGame` everywhere; remove legacy `video_game` usage in active paths).

6. Tests and docs
- Update tests to current tRPC architecture and two-table model.
- Add/adjust coverage for `collect`, `save`, `edit`, and `getPaginated` across all media types.
- Update docs that still imply the old separate non-book table architecture.

## Execution Order

1. Create schema migration (`other_media` + data migration).
2. Update backend routers/types to read from new structure.
3. Update UI contracts/components.
4. Remove dead/commented code.
5. Run lint/typecheck/tests and fix regressions.
6. Final pass on docs/tests.

## Success Criteria

- All media types function end-to-end with only two models (`books`, `other_media`).
- No endpoint accepts a type it cannot correctly process.
- UI and backend share the same sort/filter/type contract.
- Commented legacy branches are removed from active code.
- Tests validate both book and non-book paths.
