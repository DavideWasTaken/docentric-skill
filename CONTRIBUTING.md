# Contributing

Thanks for helping! The most valuable contributions right now:

1. **Designer round-trip reports** — open a template the skill generated or edited in the Docentric Designer and tell
   us what happened (issue template *Designer round-trip report*).
2. **Format knowledge** — anything that corrects or completes
   [skills/docentric/references/format.md](skills/docentric/references/format.md): other Docentric versions, ER data
   sources, AggregateFunction codes, Group elements.
3. **Expression functions** the preview does not know yet (they show up as `⟨name()⟩`).
4. **Prompts where Claude did the wrong thing with the skill** — they become test cases in
   `skills/docentric/evals/`.

## The one hard rule: no real customer data

Real Docentric templates embed **sample data copied from D365** — customer and vendor names, addresses, VAT numbers,
contract notes — and often a Microsoft Purview sensitivity label. **Never commit them, never attach them to issues.**

- Tests build synthetic templates with `tests/fixtures/builder.ts` or from specs in `examples/`.
- To reproduce a bug from a real template, rebuild the relevant part synthetically (a few elements and a tiny schema
  are usually enough).
- `private/`, `*.ddsp` and `out/` are git-ignored for local work.

## How the repository fits together

- `src/` is the tool (TypeScript). `npm run build` bundles it into `skills/docentric/scripts/docentric.mjs` — a
  generated file, committed so the skill works without an install step. CI fails if it is out of date: run
  `npm run build` and commit the result with your change.
- `skills/docentric/SKILL.md` is what Claude reads when the skill triggers; keep it short and explain *why*, details
  go to `references/`. `tests/skill.test.ts` checks the frontmatter, the links and that every edit operation is
  documented.
- `skills/docentric/evals/` holds test prompts with and without the skill (format of `claude plugin eval`).
- Validation rule codes (`DCX0nn`) are stable: add new ones at the end, never renumber, and document them in
  `references/validation-rules.md`.

## Development

```bash
npm ci
npm run typecheck
npm run lint
npm run format:check      # npm run format to fix
npm run test:run
npm run build
```

## Pull requests

Small and focused, with a test. Describe what you verified (unit tests, Word, Designer, a run of the skill).
