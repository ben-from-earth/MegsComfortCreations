# Repository conventions

Next.js 15 App Router, React 19, TypeScript, tRPC, Drizzle, Jest, Tailwind 4.

- New files use kebab-case unless Next.js, a tool, or a public URL requires another name.
- Prefer slim, fast business-logic unit tests. Do not add broad mocked UI or router suites.
- Do not extract a helper solely to make it testable.
- Prefer `satisfies` or a real shared type over `as`. Document the reason if a cast is unavoidable.
- Never pass database connection URLs through a Windows shell. Load an explicit env file with `override: true`.
- Read `ai-assistance/learnings.md` before making decisions. Append an entry there when a preventable mistake is corrected.
