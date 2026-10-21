# AGENTS.md

1. Before any reasoning, read all skill files under `.agents/skills/**/SKILL.md`, then read `ai-assistance/learnings.md`.

## Purpose

Provide shared operating instructions for AI agents working in this repository.

## Core Working Rules

- Follow repository conventions in `CLAUDE.md`.
- Keep changes scoped and minimal to the requested task.
- Run targeted validation (tests/lint/type checks) for edited areas.
- When fixing a bug or shipping a feature, add or update automated tests in the same work so behavior is protected from regression.

## Learnings Discipline

- Treat `ai-assistance/learnings.md` as required context before making decisions.
- If the user corrects agent behavior or a preventable mistake is identified, append a concise entry to `ai-assistance/learnings.md` in the existing format.
- Entries should describe: issue, correction, and the durable rule going forward.

## Notes

- This file is a policy contract for workspace agents. Compliance depends on the agent/runtime honoring repository instruction files.
