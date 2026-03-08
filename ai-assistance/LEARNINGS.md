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
