# Cleanup checklist

Work top to bottom. **Generic always applies.** Open a stack addendum only
when Step 0 matched that language in the scoped diff.

| Scoped diff              | Use              |
| ------------------------ | ---------------- |
| Any                      | Generic          |
| `*.ts`, `*.tsx`, `*.js`  | TypeScript       |
| React `*.tsx` / `*.jsx`  | React and Next   |
| CSS-only or docs-only    | Generic only     |

A React + TypeScript feature uses Generic + TypeScript + React. A CSS-only
tweak uses Generic and stops.

## Generic

### Structure

- One module, one job. A builder that also hides a second calculator, or a
  page that owns drag, chrome, and time formatting, is a candidate to split
  along the existing seam.
- Do not extract a helper for a 1–2 line repeat. Do extract a named
  function when the same rule is copied with slight drift.
- Prefer a mode or a small child over a boolean-prop pile.
- Delete dead methods, unused exports, leftover feature-flag branches,
  and comments that describe a previous design.

### Naming

Read every new or changed name (variables, parameters, functions, props,
classes, files, routes, test names) and ask whether a newcomer could say what
it holds or does without reading its body.

- Name by purpose and domain meaning, not by type or mechanism: `placedKeys`
  over `set`, `invoicePeriod` over `value`, `suppressClickAfterDrag` over
  `handler`.
- Booleans read as questions (`isLoading`, `hasChart`, `canManage`).
  Functions start with a verb that matches what they do; a getter that
  queries or computes says so (`countFor`, `forLayout`).
- Collections are plural, single items singular, maps say their key
  (`countsByWidget`).
- Drop vague stand-ins (`data`, `info`, `item`, `result`, `temp`, `x`) unless
  the scope is one or two lines and the type makes the meaning obvious.
- One concept, one word across the backend and frontend. If a tRPC router
  says `placement`, the React code should not call the same thing a `slot`
  or `entry`.
- A name that needs a comment to explain it should be renamed instead.
- Rename only within scope. Update every reference, including tests.

### Types

Too wide (causes extra logic):

- A looser declared type than the value actually is, then conversions or
  empty checks at every callsite.
- `unknown` / `any` / blanket casts that exist only to silence a missing
  field. Add the field upstream.

Too narrow (causes extra logic):

- Non-null assertions or dummy defaults for values the query already
  treats as optional.
- A stored enum forced to represent a derived filter, then special cases
  to keep it out of persistence.

Model the real shape so the guards disappear.

### Tests

- Keep coverage that pins a business rule.
- Rewrite assertions that hardcode an incidental shape the cleanup
  changed.
- Do not add tests for "the helper returns the formatted string" or
  "the filter applies when applied."

### What good enough looks like

Stop when the scoped code reads like the files next to it, the extra
conversions are gone, and a new reader can name the types without
opening three helpers. Leave residual nits in the wrap-up instead of
opening a second refactor.

## TypeScript

When `*.ts` or `*.tsx` is in scope.

- Prefer `satisfies`, existing repo types, or a new shared type over `as`.
  Use `as` only when there is no safer alternative and document why inline.
- New files use kebab-case (`media-checkboxes.tsx`) unless the name is
  required by Next.js, a tool, or a public URL. Exported symbols stay
  PascalCase or camelCase.
- Do not extract a helper solely to make it testable. Leave simple form
  wiring inline. Extract when the same logic is reused or the name itself
  is the abstraction.
- No synonym helpers or alias exports for naming convenience.
- If the scoped diff touches local DB scripts, never pass connection URLs
  through a Windows shell. Load explicit env files with `override: true`.

## React and Next

When React or Next files are in scope.

- Pages live under `app/`. Feature UI stays next to the feature
  (`app/mediacollector`, `app/showdatabase`). Shared chrome stays in
  `app/components`.
- Reuse existing tRPC hooks, form helpers, and UI primitives. Do not add
  a new context, store, or wrapper "just in case."
- Derive during render. Do not store filtered lists or formatted labels
  in state, and do not model a click as `useState` + `useEffect`.
- Index maps for repeated lookups in a list.
- Keep presentational children dumb; lift shared state to the existing
  page or form, not a new layer.
- Direct imports from source files when a barrel would pull a tree.
- When React performance or composition is in the leftover mess, follow
  `vercel-react-best-practices` and `vercel-composition-patterns`.
- Do not add aria-labels unless the user asks for them.

## Jest

When tests are in scope or the cleanup would break them.

- Prefer slim, fast business-logic unit tests.
- Avoid broad mocked UI/component suites and heavily mocked router
  suites unless that approach is already intentional in the scoped files.
- Update tests the cleanup would break. Do not add wiring, formatting,
  or "it renders" coverage.
- Do not extract a helper just so a test can import it.
