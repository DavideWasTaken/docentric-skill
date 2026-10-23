# Edit operations

`node scripts/docentric.mjs edit <template.docx> <ops.yaml> [--out <new.docx>] [--dry-run] [--force] [--preview]`

The operations are applied in order to a **copy** of the template. The result is validated; if it has errors the
original did not have, nothing is written (the message names them — fix the ops; `--force` saves anyway, for the
rare case where the user explicitly wants a binding that does not exist yet). `--dry-run` reports without writing.
Element ids come from `inspect` (`#1200000001` or `1200000001`). Ops can also be passed on stdin with `-`.

## Operations

| op | Fields | Effect |
|---|---|---|
| `insert` | `anchor` ({element: id} or {text: "…"}), `position` (after · before · inside), `blocks` | Inserts layout blocks (same syntax as the layout spec). After a field in a table cell → inside that cell. `inside` an If/List → at its end. A `text` anchor matches the design-time text of a paragraph (a field shows its path, a label its code). |
| `add_column` | `list`, `value`, `header`?, `after`? (an element in the row), `widthCm`? (exact width) or `width`? (relative to the average cell of the row, default 1), `align`?, `format`? | Adds a column to a table List: header cell (nearest row above with text), value cell, table grid. The table keeps its width: the space comes from the columns wider than average (typically the description), narrow numeric columns keep theirs — the result line says which columns shrank. New cells copy the neighbour's look. Cells spanning the insertion point grow instead. |
| `set_binding` | `element`, `path`?, `source`?, `format`? (null removes) | Changes a binding. |
| `wrap_if` | `element`, `condition`, `source`?, `level`? (auto · paragraph · row) | Shows the element only when the condition holds. auto: a field in text → its paragraph; a cell-level field → the cell; `row` → the whole table row. |
| `remove` | `element`, `keepContent`? | Removes an element and everything inside it. Table cells stay, empty, so the grid stays valid. `keepContent: true` keeps the content as static text. |
| `set_label` | `code` (with or without `@DOC_LABEL:`), `texts` {language: text} | Adds a template label or updates its texts. |
| `add_language` | `language`, `default`? | Adds a template language (then add its texts with `set_label`; `validate` lists the labels still missing it). |
| `resize_column` | `element` (any element in that column), `widthCm` | Sets the width of a table column in cm — e.g. so item codes stop wrapping. The table keeps its width: extra space comes from the wide columns, freed space goes to the widest one. Use it instead of editing `w:tblGrid` by hand. |
| `remove_column` | `element` (any element in that column) | Removes a table column: its cell in every row (header, data, totals), the elements inside them, the grid column. Cells spanning across it shrink; the freed width is spread over the other columns in proportion. |
| `set_style` | `anchor` ({element} or {text}), `bold`?, `italic`?, `underline`?, `caps`?, `size`? (pt), `color`? (RRGGBB), `align`? | Formats an element or a paragraph: runs, the content control (what the Designer shows) and, with `align`, its paragraphs. |
| `convert_placeholders` | — | Turns `{{…}}` typed directly in Word into real Docentric elements (see `layout-spec.md` for the syntax). Paragraphs with fields, links or drawings are skipped and reported. |

## Placeholders in values and blocks

`{{PurchOrderHeader/@PurchId}}` field · `{{@Qty * @PurchPrice | n2}}` with format · `{{GeneralData: CurrentCompany/@Name}}`
other data source · `{{label:Total}}` / `{{label:@SYS12345}}` label. Inside a table List, `{{@Field}}` is the current
row. Full syntax: `layout-spec.md`.

## Examples

Add the vendor's item number after the item code in the purchase order lines, with a translated header:

```yaml
- op: set_label
  code: VendItem
  texts: { en-us: Vendor item, it: Cod. fornitore }
- op: add_column
  list: "#1200000010"           # the List with RenderMode=2 that wraps the lines row
  after: "#1200000011"         # the @ItemId field in that row
  header: "{{label:VendItem}}"
  value: "{{@ExternalItemId}}"
  width: 0.8
```

Show the notes only when filled in, and add a reference line under the order number:

```yaml
- op: wrap_if
  element: "#1200000020"
  condition: "PurchOrderHeader/@Notes != ''"
- op: insert
  anchor: { element: "#1200000001" }
  position: after
  blocks:
    - text: "Ref. {{PurchOrderHeader/@VendorRef}}"
      style: { size: 8, italic: true }
```

Add German:

```yaml
- op: add_language
  language: de
- op: set_label
  code: Total
  texts: { de: Gesamt }
```

Layout too fiddly to describe (e.g. a field in a specific spot of a complex header)? Ask the user to type
`{{PurchOrderHeader/@DlvDate | d}}` where it belongs in Word and save; then run `convert_placeholders`.
