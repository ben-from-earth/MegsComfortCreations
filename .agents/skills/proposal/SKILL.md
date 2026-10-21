---
name: proposal
description: Investigate a planned implementation and write a proposal.md with findings and a segmented commit plan. Use when the user says /proposal, asks for an implementation proposal, or wants a reviewable step plan before coding.
---

# Proposal

Calling this skill requires an input of what the planned implementation is for. If that idea is missing, ask for it and stop.

The skill just takes that idea, does a full investigation, asks any necessary questions, then puts together a proposal.md outlining everything discussed, results of the codebase investigation, and explains all that and at the end puts together an implementation plan breaking up the implementation into digestable commit steps that build off of each other. No designated number of commits, just a logical amount to where each step is not too large, but small enough to properly review each step as its implemented. As you go through the steps also, we do not want to commit the files. The idea is to implement the steps, and let the user go through and review and commit manually.

## Workflow

1. Confirm the implementation idea. Refuse to invent a second idea.
2. Read matching `.ai/rules` and inspect current code, types, tests, and seeders that the idea would touch.
3. Ask only questions that would change the plan. If a default is obvious, state it in the proposal instead of blocking.
4. Write the proposal file. Do not implement the plan unless the user explicitly asks. Do not commit.

## Output path

Use the path the user names. Otherwise write `docs/<area>/proposal.md` for the feature area (`docs/schedule/proposal.md`, `docs/jobs/proposal.md`). Do not replace a Linear spec unless the user asks.

## Proposal shape

```markdown
# <Idea>

## Problem
## Outcome
## Investigation
## Decisions
## Out of scope
## Implementation plan
```

Investigation covers current behavior, the files that own it, and constraints the plan must respect. Decisions record settled choices and recommended defaults. Each implementation step is one reviewable commit: why it exists, files likely to change, what becomes true afterward, and how to verify. Later steps may depend on earlier ones; they must not bundle unrelated work.

## Schema

Edit the original migration that created the table. Do not add a follow-up migration for columns, constraints, or indexes. Name that original file in the plan. When executing a step, change the original migration and delete an alter migration added for the same change.

## Implementation later

When the user asks to execute the plan, implement one step at a time unless they ask for more. Leave git staging and commits to the user. Do not run `git commit` or `git add` as part of this skill.
