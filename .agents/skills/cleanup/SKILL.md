---
name: cleanup
description: >-
  Audit recent work and apply a focused cleanup pass. Detect stack and
  in-scope languages first, then apply matching local skills plus a generic
  quality pass. Use when the user says /cleanup, "cleanup pass", "clean this
  up", or wants leftover implementation debt removed after a feature lands.
---

# Cleanup

This skill **edits code**. It is not a review report. `/cleanup` makes the
smallest safe cleanups and leaves git to the user.

Do not commit. Do not invent a second architecture. Prefer the pattern already
used by sibling files.

## Scope

Default to the current feature: uncommitted changes plus the commits that
landed it. If the user names paths or a diff range, use that instead. Do not
wander into unrelated modules.

## Step 0: Detect

Do this before opening a skill or running a formatter. Detection is two
lists: **stack** (what the repo is) and **languages** (what this diff
touches). Activate and verify from the intersection.

### Stack (manifests, not the diff)

At each workspace root, look for files that exist — do not assume another
project's toolchain:

| Signal                                      | Stack                    |
| ------------------------------------------- | ------------------------ |
| `package.json` depends on `next`            | Next.js                  |
| `package.json` depends on `react`           | React                    |
| `package.json` depends on `drizzle-orm`     | Drizzle                  |
| `package.json` depends on `@trpc/server`    | tRPC                     |
| `jest` in `package.json` or `jest.config.*` | Jest                     |
| `typescript` / `typecheck` script           | TypeScript check         |
| `package.json` depends on `tailwindcss`     | Tailwind                 |
| `AGENTS.md` or `ai-assistance/learnings.md` | project conventions      |
| `.agents/skills/`                           | local skills to activate |

A missing signal means that tool or skill is not in this repo. Skip it.

### Languages (scoped diff)

Map the scoped paths, not the whole tree:

| Paths                                                                      | Language                |
| -------------------------------------------------------------------------- | ----------------------- |
| `*.ts`, `*.tsx`, `*.js`, `*.jsx`                                           | TypeScript / JavaScript |
| `*.tsx`, `*.jsx`, or files under a React `app` / `components` tree         | React                   |
| `*.css`                                                                    | CSS                     |
| `*.md` and other docs/config-only files                                    | docs                    |

`*.tsx` that only changes class names is still TypeScript and React, not
"docs."

### Activate

Activate a local skill only when **both** the stack exists and the language
is in the scoped diff:

| When                                                                             | Activate if the file exists                         |
| -------------------------------------------------------------------------------- | --------------------------------------------------- |
| React / Next files in scope                                                      | `vercel-react-best-practices`                       |
| React composition, boolean-prop piles, or reusable component APIs in scope       | `vercel-composition-patterns`                       |
| Tests in scope and Jest stack                                                    | this repo's slim Jest rules in [REFERENCE.md](REFERENCE.md) |

Do **not** activate a React skill because the repo has React when the scoped
diff has no React files.

Then: read `AGENTS.md`, `CLAUDE.md`, and `ai-assistance/learnings.md` when
those files exist. If `.ai/rules/index.md` exists, read every rule whose
globs match the scoped paths. Read the scoped files in full, not only hunks.
Check siblings before calling something non-idiomatic.

## Workflow

1. **Restate the leftover mess** in a few bullets (duplication, extra
   guards, baggy types, a module doing two jobs). Skip style nits.
2. **Apply the pass** using [REFERENCE.md](REFERENCE.md). Generic sections
   always, including a naming pass over every new or changed name. Stack
   addenda only when Step 0 matched that language. Each change must remove
   complexity, not relocate it. Extract only when a real boundary already
   exists or the same 5+ line rule is copied.
3. **Keep tests honest.** Update tests that the cleanup would break. Do not
   add wiring, formatting, or "it renders" coverage. Follow
   `ai-assistance/learnings.md`.
4. **Verify** with tools that exist **and** match the diff. See below.
5. **Tell the user** what changed, what you left alone and why, and what
   they should still eyeball.

## Verify

Run the smallest check that can fail for the languages you touched. If the
tool is not in the repo, skip it.

| Diff                                                 | Run                                      |
| ---------------------------------------------------- | ---------------------------------------- |
| Docs / Markdown / comments only                      | Nothing. No `tsc`, no tests.             |
| CSS only                                             | Nothing unless a CSS formatter is part of this repo's usual check. |
| JS/TS tests in scope or likely broken, Jest exists   | `npx jest` on the affected test files    |
| TypeScript sources changed and `typecheck` exists    | `npm run typecheck`                      |

Docs-only diffs do not run Jest or `typecheck`.

## Hard stops

- No drive-by refactors outside scope.
- No new folders, packages, or "service" layers.
- No synonym helpers or alias exports for naming convenience.
- No comments that restate the code.
- No extract-a-helper just to test it.
- No commit, add, or push.
