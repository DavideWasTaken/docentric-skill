---
name: docentric
description: Read, validate, preview, edit, create and convert Docentric AX report templates — the Word (.docx) templates that Dynamics 365 Finance & Operations (D365FO, AX) prints business documents with, through Docentric. Use this skill whenever the user mentions Docentric, a D365/AX document or report template in Word, a DDSP, print management templates, or report names such as PurchPurchaseOrder, SalesInvoice, SalesConfirm, CustAccountStatement; whenever they want to add, move, hide or translate a field, column, label or section in such a template, find out why a field prints empty or wrong, see what a template will look like, build a new document layout from D365 data, or convert/migrate an SSRS report (AxReport) to Docentric — even if they only attach a .docx and say "add the delivery date" without naming Docentric.
---

# Docentric templates for Dynamics 365 F&O

A Docentric template is a normal Word document plus a hidden layer: every dynamic piece (field, label, table row
repeater, condition, variable, total) is a content control tagged `<projectId>-<elementId>`, registered in a
custom XML part (`ProjectData`) together with the data source schemas (XSD + sample data exported from D365 as a
DDSP) and the template's own labels. Docentric fills it at print time.

The two layers must stay consistent, the element nesting must match the document, and every binding must exist in
the schema. Small mistakes produce a file that opens fine in Word and breaks in the Docentric Designer or at print
time. That is why this skill ships a tested tool — use it instead of editing the XML by hand.

## The tool

```bash
node <this-skill-dir>/scripts/docentric.mjs <command> …
```

`<this-skill-dir>` is the folder this SKILL.md is in. One self-contained file, Node.js 18+, nothing to install.
Check `node --version` once; if Node is missing, tell the user how to get it (Windows: `winget install
OpenJS.NodeJS.LTS`, else nodejs.org) rather than working around it.

| Command | Use it to |
|---|---|
| `inspect <t.docx> [--schema] [--labels] [--source MainData] [--depth N]` | See the report, DSP class, languages, data sources and the element tree: ids, bindings, and the **data context** of each (`.(ReportData/PurchOrderLines)` = relative to the current row). `--schema` lists every available field. |
| `validate <t.docx> [--json] [--min error]` | Offline checks DCX001–DCX019 (bindings vs schema with did-you-mean, orphans, labels, variables). Exit code 1 on errors. |
| `preview <t.docx> --formats png,pdf [--lang it] [--out dir]` | Fill the template with its sample data and write `<name>.preview.pdf/.html/.docx` and one PNG per page (`.preview.png`, `.preview.p2.png`, …). |
| `edit <t.docx> <ops.yaml\|-> [--out new.docx] [--dry-run] [--preview]` | Change a template: insert blocks, add or resize a table column, change a binding, wrap in a condition, remove, labels, languages, convert `{{…}}` typed in Word. |
| `create <spec.yaml\|-> <out.docx> [--preview]` | Build a new template from a layout spec. |
| `convert-ssrs <AxReport.xml> <out.docx>` | SSRS report → `<out>.spec.yaml` + template, with notes on what was not converted. |
| `doctor` | Which PDF/PNG converters are available (LibreOffice, Word, Edge/Chrome). |

## Workflow

1. **Understand before changing.** Run `inspect` on the template (add `--schema` when you need field names). Note
   the element ids you will touch and the data context of the area: inside a table List, bindings use `Source='.'`
   and paths like `@ItemId` are relative to the row; elsewhere they start at the data source root
   (`PurchOrderHeader/@PurchId`).
2. **Check the data exists.** The field must be in the schema (`inspect --schema`). If it is not, do not invent a
   path that "probably" exists — the template would validate as broken or print empty. Read
   `references/d365-workflows.md`: the field has to be added in X++ (the DSP class named by `inspect`) and the DDSP
   regenerated in D365. If the d365fo-mcp tools are available, use them for the X++ side.
3. **Make the change with the tool.**
   - Existing template → write the operations to a small YAML file next to the template (keeps the change reviewable
     and repeatable) and run `edit`. Operations and examples: `references/edit-operations.md`.
   - New template → write a layout spec and run `create`. Reuse the data sources of an existing template for the
     same report (`dataSources: { from: existing.docx }`) whenever one exists; a declared schema is only provisional.
     Spec reference: `references/layout-spec.md`.
   - The tool never overwrites the original: `edit` writes `<name>.edited.docx` (or `--out`), and refuses to save if
     the result has validation errors the original did not have. Read its message: it names the binding that does
     not resolve and the closest existing names.
4. **Look at the result.** Run `preview --formats png,pdf` on the new file and open the PNG of the page with the
   change yourself (line tables are often on page 2: `.preview.p2.png`) before saying it is done: check the new
   column/field is where the user asked, in the right order, with sensible values, and that headers still fit. If
   something is off, fix the ops and rerun — don't hand the user an unchecked file.
5. **Report.** What changed (element ids), the validation result, the paths of the template and preview, and what is
   left for the user: open it once in the Docentric Designer, and — for new fields — the X++/DDSP steps.

## Things that are easy to get wrong

- **Design-time text.** In a template a Field shows its binding path as text and a Label shows its code
  (`@DOC_LABEL:Total`). Text anchors for `insert` match that text, not the printed values.
- **Two kinds of labels.** `@DOC_LABEL:X` labels belong to the template (translations live in the template).
  `@SYS…`/`@ModelLabels…` are D365 labels resolved at print time; the template only caches their text. Don't
  "translate" D365 labels in the template — check them in D365 (d365fo-mcp `labels`).
- **Tables.** A List that repeats table rows wraps whole rows; Fields in tables usually wrap whole cells. Use
  `add_column`, `remove` (which keeps the cell so the grid stays valid) and `wrap_if` rather than moving XML around.
- **Previews are approximations.** They use the sample data stored in the template (often old or partial) and
  reimplement the expression functions; unknown functions show as `⟨name()⟩`. Treat the preview as a layout check;
  Docentric's own preview in D365 or the Designer is the final word.
- **Personal data.** Templates embed sample data copied from D365: customer/vendor names, addresses, VAT numbers.
  Don't paste large chunks of it into the conversation, and don't upload templates anywhere.

If the tool cannot do something the user needs (an element type it does not generate, an unusual layout), you may
edit the package directly — but read `references/format.md` first, keep the registry and the content controls in
sync, write to a new file, and run `validate` and `preview` afterwards.

## Converting SSRS reports

`convert-ssrs` reads the AxReport metadata XML (in PackagesLocalDirectory:
`<Package>/<Model>/AxReport/<ReportName>.xml`; with d365fo-mcp, `get_object_info` with objectType `report` finds
it). It writes a layout spec plus a first template and lists what it could not translate (`⟦SSRS: …⟧` markers:
custom code, group rows, sub-reports). Fix those in the spec, run `create` on it, preview, iterate. The schema it
derives from the datasets is provisional (DCX019): the user must register the template in D365, generate the DDSP
and update the data source in the Designer.

## References

- `references/edit-operations.md` — every `edit` operation with examples (read before writing ops)
- `references/layout-spec.md` — the layout spec for `create`, placeholders `{{…}}`, blocks (read before writing a spec)
- `references/validation-rules.md` — what each DCX code means and how to fix it
- `references/d365-workflows.md` — missing fields, DDSP, labels, SSRS migration, with or without d365fo-mcp
- `references/format.md` — the Docentric package format, for the rare direct edit
