# Proposal: Clearer Partial Database Save Error Dialog

Date: 2026-08-02  
Status: Implemented (2026-08-02)  
Area: Media Collector — database save + PNG export flow

---

## Summary

When **Create PNG Export** hits one or more database save failures, successes are already persisted and PNG creation is correctly cancelled. The gap is the **error dialog copy**: it is count-heavy and not tied to the visible collector blocks, so users cannot easily find and fix the failures.

This proposal keeps **Close** as a dismiss-only acknowledgment and reworks the dialog to name each failed title with its **Block #** and a **user-friendly reason**. No second save and no continue-PNG action.

---

## Current Behavior (Investigation)

### Flow today

`Create PNG Export` (`app/mediacollector/page.tsx` → `handlePNGClick`):

1. Validates every non-database block has at least one selected image.
2. Requires a PNG format selection.
3. Calls `onSubmit` → `trpc.database.save` with all `collectedData`.
4. If any result has `'error' in res`, shows the database-save dialog and **returns** (PNG is not created).
5. If all saves succeed, builds PNG images and downloads the file.

### Save semantics (already correct for this request)

`database.save` processes items independently:

- Skips `isDatabase: true` items.
- Persists each success before moving on.
- Pushes a failure result and continues for the rest of the batch.

So by the time the dialog opens, **all successes from that attempt are already saved**. No additional save step is needed on error acknowledgment.

### Current dialog copy

`InformationalDialog` (`variant="databaseSave"`) shows:

- “Results of saving to database”
- Successful / failed counts
- Nested lists: `Error saving {title} to database` + raw `errors[]` strings
- **Close** only

### Block numbering already in the UI

`CollectedCoversBlock` labels each card as `block: {index + 1}` where `index` is the 0-based position in `collectedData`. Dialog “Block #N” should use that same 1-based number so users can match the list to the grid.

### Gaps vs desired UX

| Gap                              | Why it matters                                                                               |
| -------------------------------- | -------------------------------------------------------------------------------------------- |
| No block number in the dialog    | Duplicate titles are ambiguous; users cannot jump to the right card                          |
| Error text is technical / nested | Harder to act on (“Schema Violation”, Zod tree strings, etc.)                                |
| Counts-first framing             | Desired copy leads with “[number] titles experienced errors…” and implies the rest succeeded |
| Error responses lack `blockID`   | Client can only match failures by `title` today                                              |

---

## Desired Flow (Updated Guidance)

1. User clicks **Create PNG Export**.
2. Save runs; **N** items fail for some reason.
3. Successes from that attempt are already on the happy-path save path (no extra save on error).
4. Surface the error dialog (PNG is **not** created).
5. User clicks **Close** (“okay, good to know”) — dialog dismisses; they fix issues however they see fit and can retry later.

---

## Proposed Dialog Copy

```
N titles experienced errors when attempting to save to the database.
All blocks besides the following were successfully saved:

1. [title] in Block #[block number]: [User friendly error reason]
2. ...
```

- **N** = number of failed save results.
- **title** = media title from the failure payload.
- **Block #** = 1-based index in current `collectedData` (same as `block: N` on the card).
- **User friendly error reason** = short, actionable sentence (not raw Zod trees or internal exception text).
- **Close** remains the only action; it dismisses and does **not** create a PNG.

If every attempted save failed, keep the same list pattern; the “all blocks besides the following were successfully saved” sentence still works when the “besides” set is the full attempted set (optional tighter wording for the all-failed case can be a polish pass).

---

## Proposed Solution

### Control flow

Keep the existing early return after partial/total save failure. Do **not** add a continue-PNG button and do **not** call `database.save` again from the dialog.

**In this pass:** after a partial save (when opening the error dialog), mark every successfully saved block `isDatabase: true` in form state. Use success payloads’ top-level `blockID` (`success: true` results) to update matching `collectedData` entries. That way a later **Create PNG Export** skips those blocks in `database.save` and will not create duplicates. Failed blocks stay `isDatabase: false` so they can be fixed and retried.

`database.save` items are a `success`-discriminated union (`DatabaseSaveResultItem`) with top-level `blockID` + `title` on both branches. Edit responses stay on the older shapes for now.

### Match failures to Block

Add `blockID` to `DatabaseSaveEditErrorResponse` and populate it on every save error path.

On the client:

```ts
const failedItems = databaseSavedData.filter((item) => 'error' in item);

const lines = failedItems.map((item) => {
  const blockIndex = collectedData.findIndex(
    (block) => block.blockID === item.blockID,
  );
  const blockNumber = blockIndex >= 0 ? blockIndex + 1 : null;
  return { title: item.title, blockNumber, reason: toUserFriendlyReason(item) };
});
```

Highlight failed cards with existing `blockIdsWithErrors` using those `blockID`s when the dialog opens (same visual cue used for missing-image validation).

### User-friendly error reasons

Map known server error kinds to stable copy for the dialog. Keep technical detail in logs/server payloads if needed, but show one sentence in the UI.

Suggested starting map:

| Server `error` / situation                       | User-facing reason                                                                              |
| ------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| `Image Persistence Error` (creation rolled back) | The cover image could not be saved, so this item was not added to the database.                 |
| `Schema Violation`                               | Some required details for this item were missing or invalid.                                    |
| `Database Insertion Error` with genre missing    | A selected genre is not available in the database.                                              |
| Other `Database Insertion Error`                 | This item could not be saved to the database. Try again, or remove the block and re-collect it. |
| Unknown / fallback                               | This item could not be saved to the database.                                                   |

Prefer mapping from the structured `error` field (and maybe a small set of known `errors[]` prefixes) rather than dumping `errors.join(...)`.

### Dialog UI changes

Update `InformationalDialog` `databaseSave` variant:

- Replace count + nested lists with the copy above.
- Keep a single **Close** button.
- Pass enough props to render Block # and friendly reasons (either preformatted lines from the page, or `data` + `collectedData` / display models).

---

## Implementation Plan

1. **Types / API** — Add `blockID` to `DatabaseSaveEditErrorResponse`; set it on all `database.save` error pushes; update router tests.
2. **Friendly reason helper** — Pure function from error response → display string; unit test the map.
3. **Dialog copy** — Rework `InformationalDialog` database-save layout to the new wording and numbered failure lines.
4. **Page wiring** — When opening the dialog after partial failure, resolve Block # from `collectedData`, set `blockIdsWithErrors`, and mark successful blocks `isDatabase: true` via success `actionAttemptItem.blockID`.
5. **Tests** — Router `blockID` on errors; dialog renders N / title / Block # / friendly reason; Close does not trigger PNG; successful blocks flip to `isDatabase: true` after a partial save.

---

## Acceptance Criteria

- [ ] On partial save failure during **Create PNG Export**, PNG is not created.
- [ ] Successes from that save attempt remain saved (existing server behavior).
- [ ] Dialog copy matches the approved pattern (N titles… / list with title, Block #, friendly reason).
- [ ] Block # matches the on-card `block: N` numbering.
- [ ] **Close** dismisses the dialog and does not create a PNG.
- [ ] Failed error responses include `blockID`.
- [ ] Failed blocks are highlightable via existing error styling.
- [ ] Successfully saved blocks are marked `isDatabase: true` in form state when the error dialog opens; failed blocks are not.
- [ ] A later **Create PNG Export** after Close does not re-insert those successes.
- [ ] Automated tests cover dialog rendering, error `blockID` plumbing, and the `isDatabase` flip.

---

## Out of Scope

- **Create PNG anyways** / continue export despite failures (explicitly dropped in this revision)
- Retry-save from the dialog
- Changing per-item save semantics on the server
- PNG layout / format changes

---

## Open Questions

1. Exact button label: keep **Close**, or rename to **OK** / **Got it**?
2. For the all-failed case, keep the same “All blocks besides the following were successfully saved” sentence, or use alternate copy when success count is 0?
