# Testing the skill from scratch

A checklist for a clean machine — a Windows VM is the most realistic, since that is where D365 developers work.
About 30 minutes. Each step says what you should see.

## 1. Prerequisites

| | Needed for | Check |
|---|---|---|
| Node.js **18 or later** | the skill's tool | `node --version` (Windows: `winget install OpenJS.NodeJS.LTS`, then a new terminal) |
| Claude Code, or a claude.ai account with code execution | using the skill | `claude --version` |
| Git | cloning (optional for the plugin route) | `git --version` |
| Microsoft Word *or* LibreOffice | PDF previews (optional) | — |
| Edge or Chrome | PNG previews (optional; Edge ships with Windows) | — |

## 2. The tool alone (no Claude involved)

```powershell
git clone https://github.com/DavideWasTaken/docentric-skill.git
cd docentric-skill
$tool = "skills/docentric/scripts/docentric.mjs"
node $tool doctor
```

`doctor` lists what is available, e.g. `✔ Node.js`, `✔ Word`, `✔ Browser`. No `npm install` is needed: the script is
self-contained.

| # | Command | Expect |
|---|---|---|
| 1 | `node $tool inspect examples/SalesInvoice.docx` | Report `SalesInvoice.Report`, 33 elements; the lines table shows `.(ReportData/InvoiceLines)`. |
| 2 | `node $tool validate examples/SalesInvoice.docx` | 0 errors, 2 warnings DCX019 (provisional schema — expected for the example). |
| 3 | `node $tool preview examples/SalesInvoice.docx --out out --formats pdf,png --lang it` | `out/SalesInvoice.preview.png/.pdf` in Italian: "Fattura", amounts like `1.992,00`. |
| 4 | `node $tool edit examples/SalesInvoice.docx examples/SalesInvoice.edit.yaml --out out/Edited.docx --preview` | Three operations, "Saved", a preview with a "Disc. %" column. |
| 5 | `node $tool create examples/sales-invoice.yaml out/Created.docx --preview` | "Created … 33 elements". Open it in Word: content controls are visible (Developer tab → Design Mode). |
| 6 | `node $tool convert-ssrs examples/ssrs/ConInvoice.AxReport.xml out/ConInvoice.docx` | A `.spec.yaml`, a template, a note about a dropped group row and one ⟦SSRS⟧ expression. |

## 3. The skill in Claude Code

Install it as a plugin (inside Claude Code):

```
/plugin marketplace add DavideWasTaken/docentric-skill
/plugin install docentric@docentric-skill
```

Restart Claude Code if asked. In a folder with a copy of `examples/SalesInvoice.docx`, try in order — you should not
need to mention the skill, it triggers on its own:

1. *"What does SalesInvoice.docx print in the lines table?"* — expect `inspect`, an answer naming the columns.
2. *"Show me how it looks in Italian."* — expect `preview --lang it` and Claude reading the PNG.
3. *"Add a column with the discount percentage after the description, right-aligned."* — expect an ops file, `edit`,
   0 new errors, a preview with the new column.
4. *"Add the customer account number under the customer name."* — the field is not in the data: expect Claude to
   say so and explain the X++ / DDSP steps, **not** to bind an invented field.
5. *"Create a delivery note from the same data: company, customer, items without prices, a signature line."* —
   expect a spec, `create`, a preview.
6. *"Convert examples/ssrs/ConInvoice.AxReport.xml to Docentric and fix what you could not convert."*
7. With d365fo-mcp connected as well: *"Add the vendor item number to the purchase order lines; if it is not in the
   data, tell me what to change in X++."*

## 4. The skill in claude.ai

Download `docentric-skill.zip` from the latest release → *Settings → Capabilities → Skills → Upload skill*. Start a
chat, attach `examples/SalesInvoice.docx`, and repeat prompts 1–5. Previews there use LibreOffice in the sandbox.

## 5. With a real template (stays on the VM)

Put a copy in `private/` (git-ignored) and ask Claude to validate and preview it, or run the tool directly. Compare
the preview with what D365 prints: differences in **values** are expected where the template uses functions the
preview does not implement (`⟨name()⟩`, listed); differences in **structure** (missing rows, wrong conditions) are
bugs worth reporting — without attaching the template.

## 6. The Docentric Designer round-trip (most important)

Open `out/Created.docx`, `out/Edited.docx` and `out/ConInvoice.docx` in Word **with the Docentric Designer add-in**.
Check that the template loads without errors, the element tree shows the elements with their bindings, and the
Designer's preview works with the embedded sample data. For an edited real template, check that the existing elements
are untouched.

Report the result with the issue template **"Designer round-trip report"** — the single most useful feedback for
this project.
