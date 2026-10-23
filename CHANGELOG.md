# Changelog

## 0.3.0 — 2026-10-23

**docentric is now a Claude skill.** The earlier MCP server (docentric-mcp) is retired: the same engine ships as a
skill that Claude Code, claude.ai and Claude Desktop load on demand, with no server to configure.

- `skills/docentric/`: `SKILL.md` (when and how to use it, the pitfalls of the format), references (edit operations,
  layout spec, validation rules, D365 workflows, the package format) and `scripts/docentric.mjs`, the whole tool in
  one file for Node.js 18+.
- Claude Code plugin + marketplace (`/plugin marketplace add DavideWasTaken/docentric-skill`), zip for claude.ai.
- New edit operations `add_language` and `resize_column`; `add_column` takes an exact `widthCm`; `create` and `edit`
  read the spec/operations from stdin with `-`.
- Found by running the skill against real templates: `add_column` now takes the space from the widest column first
  (it used to shrink every column and make headers wrap); previews follow the page size and orientation and produce
  one PNG per page (a lines table is often on page 2), with the height measured by the browser.
- Test cases with and without the skill in `skills/docentric/evals/`; results in the README.
- Environment variables renamed: `DOCENTRIC_ALLOWED_ROOTS`, `DOCENTRIC_SOFFICE`, `DOCENTRIC_BROWSER`.

## 0.2.0 — 2026-10-23

- Preview (sample data → docx/html/pdf/png), create from a layout spec, edit operations, SSRS conversion, CLI,
  validation rule DCX019.

## 0.1.0 — 2026-10-23

- Read-only analysis: inspect and validate (18 rules).
