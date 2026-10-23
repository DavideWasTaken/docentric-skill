# Working with D365: fields, DDSP, labels, SSRS

The template only knows the data its data sources describe. Everything upstream — the X++ that fills the data, the
DDSP, the report registration — lives in D365. When the d365fo-mcp tools are available
([d365fo-mcp-server](https://github.com/dynamics365ninja/d365fo-mcp-server)), use them for that side; otherwise,
tell the user precisely what to do in Visual Studio / D365.

## Where the data comes from

`inspect` prints the report id (e.g. `PurchPurchaseOrder.Report`) and the **data source provider (DSP) class**
(e.g. `ConPurchPurchaseOrderDSP`, a Docentric class that usually extends `DocDataSourceProviderSrsReport` and adds
fields on top of the SSRS report's RDP temp tables). `MainData` mirrors those temp tables / DSP output; `GeneralData`
is company and user info Docentric adds; `Parameters` are the report contract parameters.

## Add a field that is already in the data

1. `inspect --schema` (or `--schema --source MainData --depth 3`) → find the path, e.g. `PurchOrderHeader/@DlvDate`.
2. `edit` with `insert` / `add_column` and `{{PurchOrderHeader/@DlvDate | d}}`.
3. `preview` and look.

## Add a field that is not in the data yet

Never bind a guessed path: it fails validation and prints empty. Instead:

1. Confirm it is missing (`inspect --schema`, and `validate` names the closest existing fields).
2. Find where it would come from: the DSP class and the temp tables of the report's RDP. With d365fo-mcp:
   `get_object_info` (objectType `class`) on the DSP class, (`table`) on the temp tables. Without: name them to the
   user from the `inspect` output.
3. The X++ change (done by the developer, or with d365fo-mcp): add the field to the temp table (or a table
   extension) and fill it in the RDP / DSP, or add it in the DSP's data-building methods. Build.
4. In D365: Docentric → generate the DDSP for the report; in the Docentric Designer, update the template's data
   source with the new DDSP.
5. Back here: bind the field with `edit`, `validate`, `preview`.

Alternatively, bind it now against a provisional schema: create the template with a declared schema that includes
the new field (rule DCX019 will remind everyone it is provisional) — useful to agree on the layout before the X++
work is done.

## Labels

- `@DOC_LABEL:X` — template labels: edit them with `set_label`, add languages with `add_language`.
- `@SYS…`, `@ModelLabels:…` — D365 labels, resolved at print time in the user's language; the template only caches
  their text. `validate` lists them (DCX011). To check or change them, use D365 (d365fo-mcp `labels`, or the label
  files in Visual Studio).

## Convert an SSRS report

1. Locate the AxReport XML: `PackagesLocalDirectory/<Package>/<Model>/AxReport/<Report>.xml` (with d365fo-mcp,
   `get_object_info` with objectType `report`).
2. `convert-ssrs <AxReport.xml> <out.docx>` → `<out>.spec.yaml`, a template, and notes.
3. Fix the notes and every `⟦SSRS: …⟧` marker in the spec (custom code, group rows, sub-reports, `Globals!…`), then
   `create <out>.spec.yaml <out>.docx --preview` and iterate on the PNG.
4. The schema is derived from the report datasets (provisional). The element names Docentric uses may differ from
   the dataset names: after generating the DDSP in D365, update the data source in the Designer and run `validate`
   — DCX007 findings with did-you-mean suggestions show what to rename (`set_binding`).
5. In D365: set up the Docentric report/template for the print destination.

## Why a field prints empty

`validate` first (DCX007 = the path does not exist). If the path exists: check the data context in `inspect`
(a field outside the table List using `@Field` reads the root, not the row), the condition of any enclosing If, and
whether the DSP actually fills the field for that document (sample data in the preview shows what the DDSP had).
