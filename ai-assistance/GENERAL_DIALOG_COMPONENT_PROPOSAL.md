# Proposal: General Dialog Component

Date: 2026-08-08  
Status: Implemented — shared `Dialog` shell + Media Collector migration  
Area: Shared UI — `app/components/ui`

---

## Summary

Introduce a reusable **Dialog** shell so feature screens can open modal content without reinventing layout, chrome, or close behavior. The shell owns positioning, styling, a titled header with a required close action, a padded body for arbitrary children, and a full-viewport backdrop that slightly blurs the page behind and blocks interaction with it.

`InformationalDialog` today is media-collector-specific and not general. This proposal extracts the shared shell first; informational / database-save content can later compose on top of it.

---

## Current Behavior (Investigation)

### What exists today

`app/mediacollector/InformationalDialog.tsx` is the only dialog-like UI in the app. It:

- Hard-codes two variants: `databaseSave` and `informationalOnly`
- Owns both shell chrome (fixed center panel, pink border/background, overflow) and feature copy
- Duplicates nearly the same outer markup in both branches
- Exposes `close: () => void` (not `onClose`) and has **no title prop**
- Uses a primary `Button` labeled `"Close"` in the header row

Shared UI under `app/components/ui` currently has `Button` and `command` only — no general Dialog.

### Gaps vs desired API

| Gap | Why it matters |
| --- | --- |
| Variant-driven content inside the shell | Every new dialog needs another variant or a fork of the component |
| No `title` on the header | Call sites cannot name the dialog without baking copy into the shell |
| Shell and feature content are coupled | Media Collector concerns leak into what should be shared UI |
| Close wiring is informal (`close`) | Inconsistent with React event naming (`onClose`) used elsewhere |
| Padding / layout differ per variant | Hard to keep a single visual standard for body content |
| No backdrop / page still interactive | Users can click through the dialog into the page behind |

---

## Desired API

Single exported component under `app/components/ui/Dialog.tsx`.

### `Dialog`

Outer shell: fixed centered panel, existing visual language (border, background, z-index, max height/width, overflow). Call sites pass title, dismiss handler, and body content.

When open, `Dialog` also renders a **full-viewport backdrop** behind the panel that:

- Slightly blurs the screen behind (e.g. `backdrop-blur-sm` plus a light translucent scrim)
- Covers the full viewport so the page underneath is **unclickable** (`pointer-events` blocked by the overlay)
- Sits under the dialog panel in stacking order (backdrop lower z-index than the panel, both above page content)

Dismiss remains header **Close** only in this pass — clicking the backdrop does **not** call `onClose`.

```tsx
interface DialogProps {
  title: string;
  onClose: () => void;
  children: React.ReactNode;
  className?: string; // optional escape hatch via twMerge
}
```

### Internal header (not exported)

`DialogHeader` is a **local helper inside `Dialog.tsx`**, not part of the public API. It renders the title on the leading side and the close control trailing (`ml-auto` / space-between), using the existing close affordance from `InformationalDialog` (primary `Button` with label `"Close"`, or migrate to `Button variant="close"` if we decide that matches better during implementation).

`Dialog` passes `title` and `onClose` down to that local header. **`onClose` is required** so every dialog has an explicit dismiss path.

### Body (`children`)

`children` are the dialog body. Apply consistent body padding (today’s useful pattern is roughly `px-10` / similar horizontal padding used in the database-save branch) so call sites do not re-specify shell spacing.

Suggested usage:

```tsx
<Dialog title="Save errors" onClose={() => setOpen(false)}>
  <p>...</p>
</Dialog>
```

---

## Example Usage After Adoption

### Informational-only (Media Collector)

```tsx
<Dialog
  title="Notice"
  onClose={() => setShowInformationalDialog(false)}
>
  <p>{informationalDialogText}</p>
</Dialog>
```

### Database-save failures

```tsx
<Dialog
  title="Database save"
  onClose={() => setDatabaseSaved(false)}
>
  <p>{failedCount} titles experienced errors...</p>
  <ol>...</ol>
</Dialog>
```

`InformationalDialog` can remain as a thin media-collector wrapper that renders those bodies inside the shared shell, or call sites can use `Dialog` directly and delete the wrapper once both variants are migrated.

---

## Implementation Plan

1. **Add shared shell** — Done: `app/components/ui/Dialog.tsx` with local `DialogHeader`.
2. **Wire props** — Done: `title`, `onClose`, `children`, optional `className`.
3. **Backdrop** — Done: full-viewport blur + scrim; blocks pointer events; does not dismiss on click.
4. **Body padding** — Done: padded body region around `children`.
5. **Migrate Media Collector** — Done: page uses shared `Dialog`; database-save body lives in `DatabaseSaveFailureBody`; `InformationalDialog` removed.

---

## Non-Goals (This Pass)

- Focus trap / Escape-to-close (can layer later)
- Dismiss-on-backdrop-click (backdrop blocks interaction only; Close remains the dismiss path)
- Portal / Radix Dialog dependency
- Multiple button actions in the footer (OK / Cancel patterns)
- Changing database-save or informational copy
- Animated open/close
- Exporting `DialogHeader` as a public compound API

---

## Open Questions

1. Keep the labeled `"Close"` primary button, or switch the header control to `Button variant="close"` (icon) for consistency with other dismiss UIs like genre chips?
2. Migrate `InformationalDialog` in the same PR as the shell, or land the shell first and migrate call sites next?
3. How strong should the blur/scrim be (subtle `backdrop-blur-sm` vs stronger), and should the scrim tint match brand pink or stay neutral?

---

## Acceptance Criteria

- [x] Shared `Dialog` lives under `app/components/ui`
- [x] `Dialog` accepts `title: string`, required `onClose: () => void`, and `children`
- [x] Header (title + close) is implemented locally inside `Dialog`, not exported
- [x] Close button matches current dismiss behavior (calls `onClose` once)
- [x] Children render in a padded body region inside the shell
- [x] Full-viewport backdrop slightly blurs the page behind and blocks clicks through to it
- [x] Backdrop click does not dismiss the dialog
- [x] Media Collector dialogs can use the shell without duplicating fixed/center chrome
- [x] Automated tests cover title, close, children rendering, and backdrop blocking interaction
