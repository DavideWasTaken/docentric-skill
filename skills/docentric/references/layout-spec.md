# Layout spec

`node scripts/docentric.mjs create <spec.yaml> <out.docx> [--preview]` builds a Docentric template from a layout
spec (YAML or JSON; `-` reads it from stdin). The spec describes **structure** — paragraphs, tables bound to
collections, conditions — and Word does the typesetting. Paths in the spec (`dataSources.from`, `baseDocument`) are
relative to the spec file. A complete, working example: `examples/sales-invoice.yaml` in the skill's repository.

## Top level

| Key | Required | Meaning |
|---|---|---|
| `dataSources` | yes | `{ from: "existing-template.docx" }` copies the data sources (schema + sample data) of a template for the same report — the reliable way. `{ schema: { MainData: {...} } }` declares them (provisional, see below). |
| `body` | yes | List of blocks. |
| `header`, `footer` | | Blocks for the page header / footer. |
| `labels` | | Template labels: `{ Total: { en-us: Total, it: Totale } }` → `@DOC_LABEL:Total`. |
| `languages` | | Template languages, first = default. Default: from `dataSources.from`, else `[en-us]`. |
| `report` | | `{ id, dspClass, group }` — SSRS report id (e.g. `SalesInvoice.Report`) and DSP class. |
| `defaultSource` | | Data source for placeholders without one. Default `MainData`. |
| `page` | | `{ size: A4\|Letter\|A5, orientation: portrait\|landscape, margins: 18 }` (mm, or [top,right,bottom,left]). |
| `font` | | `{ name: Calibri, size: 10 }` |
| `baseDocument` | | A .docx whose styles, theme and fonts are reused (corporate look). |
| `title` | | Document title. |

## Placeholders — in any text

| Placeholder | Becomes |
|---|---|
| `{{PurchOrderHeader/@PurchId}}` | Field on the default data source (or the current row inside a table/list) |
| `{{GeneralData: CurrentCompany/@Name}}` | Field on another data source |
| `{{@Qty * @PurchPrice \| n2}}` | Field with a .NET format (n2, d, dd/MM/yyyy, #,##0.00 …) |
| `{{label:Total}}` / `{{label:@SYS12345}}` | Label (template label / D365 label) |
| `{{sum: @Qty * @PurchPrice \| n2}}` | SummaryField — only in a table's `totals` |

Expressions are Docentric/XPath: `concat()`, `iif(cond, a, b)`, `format()`, `round-number()`, `upper-case()`,
comparisons (`=`, `!=`, `>`), `and`/`or`, `div`, `$new-line`. Inside a table or list, `@Field` is the current row.

## Blocks

```yaml
- text: "Order {{PurchOrderHeader/@PurchId}}"      # paragraph
  style: { bold: true, size: 12, color: "1F4E79", align: right, italic: true, caps: true, spaceAfter: 6 }
- heading: "{{label:PurchaseOrder}}"               # level 1-3
  level: 1
- grid:                                             # static table (key/value blocks, layout)
    - ["{{label:Date}}", "{{PurchOrderHeader/@OrderDate | d}}"]
    - [{ text: "{{label:Total}}", span: 2, style: { bold: true }, shade: "F2F2F2" }]
  widths: [1, 2]                                    # relative
  borders: false
- columns:                                          # side-by-side blocks, borderless
    - [ { text: "left column" } ]
    - [ { text: "right column", style: { align: right } } ]
  widths: [3, 2]
- table:                                            # data table: one row per record
    rows: PurchOrderLines                           # collection path
    if: "@IsWithAmount"                             # optional: show the table only if…
    columns:
      - { header: "{{label:Item}}", value: "{{@ItemId}}", width: 2 }
      - { header: "{{label:Qty}}", value: "{{@Qty}}", width: 1, align: right, format: n2 }
    totals:
      - [ { text: "{{label:Total}}", span: 1 }, "{{sum: @Qty | n2}}" ]
    headerShade: "F2F2F2"
- if: "PurchOrderHeader/@Notes != ''"               # conditional block
  then: [ { text: "{{PurchOrderHeader/@Notes}}" } ]
- list: PurchOrderLines                             # repeat blocks per record (not a table)
  body: [ { text: "{{@ItemId}} — {{@Name}}" } ]
- var: total                                        # Docentric variable, read with var-element-value('total')
  value: "0"
- image: PurchOrderHeader/@CompanyLogo              # Image element (mm)
  width: 40
  height: 15
  align: right
- spacer: 6                                         # empty paragraph (pt after)
- pageBreak: true
```

## Provisional schemas

Without a template to copy data sources from, declare them:

```yaml
dataSources:
  schema:
    MainData:
      root: ReportData
      attributes: { IsWithAmount: boolean }
      elements:
        PurchOrderHeader: { attributes: { PurchId: string, OrderDate: date } }
        PurchOrderLines: { repeated: true, attributes: { ItemId: string, Qty: decimal } }
```

Types: string, decimal, int, date, dateTime, boolean, base64Binary. Sample data is generated so the preview works;
to make the preview realistic, give each data source a `sample:` with XML for its root element.
The template is marked provisional (validation rule DCX019) until the real DDSP generated in D365 is loaded in the
Designer. Take the field list from the DSP's temp tables (in Visual Studio, or with d365fo-mcp `get_object_info`
on the table) rather than guessing names.

## Tips

- Prefer `dataSources: { from: … }` with an existing template for the same report: real schema, real sample data.
- Keep headers short (2–3 lines); put addresses in `columns` so they sit side by side.
- `{{sum: …}}` is the only aggregate generated (sum); for counts use a Field such as `{{count(PurchOrderLines)}}`.
- After `create`, always `preview` and look at the PNG; adjust widths (relative numbers) and styles, re-run.
