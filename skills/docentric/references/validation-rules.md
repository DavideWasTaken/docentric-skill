# Validation rules

`validate` checks a template offline. Errors break printing or render wrong; warnings are probably mistakes; info is
for awareness. `edit` refuses to save a result with new errors.

| Code | Severity | Meaning | Usual fix |
|---|---|---|---|
| DCX001 | error | A content control references an element that is not in ProjectData (often content pasted from another template). | `remove` it, or re-insert the element. |
| DCX002 | warning | An element is registered but not placed in the document (leftover from deleted content). | Harmless; clean up in the Designer. |
| DCX003 | error | `ChildElements` lists an element that does not exist. | Usually a hand edit gone wrong; re-create the element. |
| DCX004 | warning | A content control's title differs from its element type. | Cosmetic; the Designer uses the type. |
| DCX005 | error | A binding names a data source that does not exist. | `set_binding` with `source: MainData` (or the right one). |
| DCX006 | error | `Source='.'` outside any List/Group — there is no current row. | Use an absolute path with a data source, or move the element into the List. |
| DCX007 | error | A path does not exist in the data source schema. The hint lists the closest names. | Fix the path (`set_binding`), or add the field in D365 (see `d365-workflows.md`). |
| DCX008 | error | A `@DOC_LABEL:` label is used but not defined. The hint suggests similar codes. | `set_label`, or fix the code. |
| DCX009 | warning | A template label has no text for one of the template languages. | `set_label` with the missing language. |
| DCX010 | info | Template labels never used (one aggregated finding). | Optional clean-up in the Designer. |
| DCX011 | info | D365 labels used by the template — resolved at print time, not checkable offline. | Check them in D365 (d365fo-mcp `labels`). |
| DCX012 | error | `var-element-value('x')` reads a variable no Var element defines. | Fix the name or add the Var. |
| DCX013 | warning | The same element is placed twice. | Remove one copy; editing one changes both. |
| DCX014 | warning | Content controls carry more than one project id (pasted from another template). | Check those elements in the Designer. |
| DCX015 | warning | An element that needs a binding has none. | `set_binding` or `remove`. |
| DCX016 | error | A binding expression cannot be parsed. | Fix quotes/parentheses with `set_binding`. |
| DCX017 | warning | A List/Group is bound to an expression instead of a plain collection path; its rows cannot be checked. | Bind the List to the collection (`PurchOrderLines`). |
| DCX018 | warning | A data source has no XSD; its bindings cannot be checked. | Update the data source from a DDSP in the Designer. |
| DCX019 | warning | A data source schema is provisional (declared in a spec or derived from SSRS, not exported by D365). | Generate the DDSP in D365 and update the data source in the Designer before production. |
