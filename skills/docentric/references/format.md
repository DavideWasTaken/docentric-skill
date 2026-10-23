# The Docentric template format — what we know

Docentric does not document its template format. This page records what the skill's tool relies on, from reading
templates written by **Docentric Designer 6.1** (AX/D365 F&O edition). Each item says how sure we are:
**seen** (observed in real templates), **inferred** (consistent with everything seen, not confirmed), **unknown**.
Corrections from people with other versions are very welcome — open an issue with the *template not read correctly*
template.

## Package

A Word package (`.docx`) with one extra custom XML part — **seen**:

```
customXml/itemN.xml        root <ProjectData xmlns="http://Docentric.org/ProjectData.xml" ProductVersion="6.1.62.1715">
customXml/itemPropsN.xml   schemaRefs: ProjectData.xml, "", …/dynamics-ax/templates, XMLSchema, …/xml-schema-extensions
```

The part is linked from `word/_rels/document.xml.rels` (`…/relationships/customXml`).

## ProjectData

```
ProjectData
├─ UISettings                         Designer pane layout (ignored)
├─ dataSourcePackage  (ns …/dynamics-ax/templates, attributes prefixed, e.g. d2p1:)
│  ├─ ssrsReportInfo  reportId, reportGroupId, dataSourceProviderClassName, dataSourceType (SSRS/ER), er*Guid
│  ├─ languages       defaultLanguage; <language value="it"/>…
│  └─ labels          <label name code section="Custom"><text language value/></label>
├─ DataSources
│  └─ DataSource Name=MainData|GeneralData|Parameters  dataSectionType=1|2|3
│     ├─ Schema/XsdList/MainXsd/xs:schema      the XSD from the DDSP
│     └─ PreviewData/SampleXml/<root …>        the sample data from the DDSP
└─ Elements
   └─ Element Type Id [Name] [attributes]
      ├─ ChildElements/ElementRef Id         nesting (List/If/Group → their content)
      ├─ Value LabelCode                      Label elements
      ├─ ValueBinding DataKind=2 Source Path ResultDataType
      └─ AggregateBinding Path               SummaryField
```

- **seen** — The `labels` block holds both the template's own labels (`@DOC_LABEL:X`, `section="Custom"`, with a
  name) and cached texts of the D365/ISV labels used by the schema (`@SYS…`, `@DocentricAX:…`, no name, no section).
- **seen** — `Source` is a data source name, or `.` for the current item of the enclosing List (or the collection of
  a SummaryField). Paths with a data source start at its root element (`ReportData`, `GeneralData`, `Parameters`).
- **seen** — Expressions are XPath 1.0-like with Docentric functions: `iif`, `concat`, `format`, `round-number`,
  `upper-case`, `starts-with`, `replace`, `char`, `label("@DOC_LABEL:X")`, `var-element-value('name')`, variables
  `$new-line`/`$lf`, operators `div`, `and`, `or`, `!=`.
- **seen** — Element attributes: `RenderMode` on List (`2` when the List repeats table rows, `1` otherwise — the
  meaning of `1` is **inferred**), `FormatString` (.NET format), `Invert` on If, `SizeMode` on Image,
  `AggregateFunction` + `IgnoreNullValues` on SummaryField.
- **seen** — `AggregateFunction="1"` is used for a sum. The other codes (Average, Count, First, Last, Max, Min,
  StdDev per the API docs) are **unknown**, so the tool only generates `1`.
- **inferred** — `ResultDataType` values: `System.String` (Field), `System.Boolean` (If), `System.Double`
  (aggregates), `System.Byte[]` (Image).

## Content controls

- **seen** — One `w:sdt` per element: `w:alias` = element type, `w:tag` = `<projectId>-<elementId>` (both decimal),
  `w15:color` by type: Label `FFCC00`, If `0000FF`, List `FF00FF`, Var `008000`; Field and SummaryField have none.
- **seen** — The project id is not stored in ProjectData; it is the same on all tags of a template. Content pasted
  from another template keeps its project id (rule DCX014).
- **seen** — Levels: Fields and Labels in tables usually wrap the **whole cell** (`w:sdtContent/w:tc`); Lists that
  repeat rows wrap **rows** (`w:sdtContent/w:tr`), sometimes through an untagged repeating-section `w:sdt`; Ifs wrap
  cells, rows or paragraphs; Fields in text are run-level.
- **seen** — At design time a Field shows its binding path as text, a Label its code.
- **inferred** — `w:placeholder/w:docPart` and `w:showingPlcHdr` are cosmetic; the tool does not write them.

## Elements the tool does not generate yet

Group, Chart, Barcode, Hyperlink, ER-based data sources. They are read, validated and kept untouched by edits.

## Editing the package directly (when the tool cannot)

1. Unzip to a working folder; never touch the original file.
2. Element ids are positive integers unique in `Elements`; content control `w:id` values are signed 32-bit integers
   unique in the document. Tag new controls `<projectId>-<elementId>` with the template's existing project id.
3. Register every new element in `Elements` and put its id in the `ChildElements` of the element whose content
   control contains it (List/If/Group); top-level elements have no parent.
4. A table cell must keep at least one paragraph; never delete a `w:tc` of a row unless you fix `w:gridSpan` and
   `w:tblGrid` for the whole table.
5. Zip back with the same part names, then run `validate` and `preview` on the result.
