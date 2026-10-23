/**
 * Regenerates the committed examples (all synthetic):
 *   examples/SalesInvoice.docx           from examples/sales-invoice.yaml
 *   examples/SalesInvoice.edit.yaml       edit operations targeting that template's element ids
 *   examples/ssrs/ConInvoice.AxReport.xml a fictitious SSRS report for convert-ssrs
 *   docs/assets/*.png                      README screenshots (needs Edge/Chrome)
 *
 *   npm run examples
 */
import { copyFile, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parse, stringify } from 'yaml';
import { createTemplate } from '../src/create/createTemplate.js';
import { loadTemplateFile } from '../src/docx/loadTemplate.js';
import { editTemplate } from '../src/edit/editTemplate.js';
import { renderPreview } from '../src/preview/renderPreview.js';
import { convertAxReport } from '../src/ssrs/convert.js';
import { AX_REPORT } from '../tests/fixtures/axreport.js';

const spec = parse(await readFile('examples/sales-invoice.yaml', 'utf8'));
const created = await createTemplate(spec, 'examples/SalesInvoice.docx', { baseDir: 'examples' });
console.log(`examples/SalesInvoice.docx — ${created.elements} elements, ${created.validation.counts.error} errors`);

const t = await loadTemplateFile(created.path);
const list = [...t.elements.values()].find((e) => e.type === 'List' && e.binding?.path === 'InvoiceLines');
const nameField = list?.childIds.map((id) => t.elements.get(id)).find((e) => e?.binding?.path === '@Name');
const notes = [...t.elements.values()].find((e) => e.type === 'Field' && e.binding?.path === 'InvoiceHeader/@Notes');
const ops = [
  { op: 'set_label', code: 'Discount', texts: { 'en-us': 'Disc. %', it: 'Sconto %' } },
  {
    op: 'add_column',
    list: `#${list?.id}`,
    after: `#${nameField?.id}`,
    header: '{{label:Discount}}',
    value: '{{@DiscPercent}}',
    width: 0.8,
    align: 'right',
    format: 'n1',
  },
  {
    op: 'insert',
    anchor: { element: `#${notes?.id}` },
    position: 'after',
    blocks: [
      { text: 'Thank you for your business, {{InvoiceHeader/@CustName}}.', style: { italic: true, color: '595959' } },
    ],
  },
];
await writeFile(
  'examples/SalesInvoice.edit.yaml',
  `# Edit operations for examples/SalesInvoice.docx (element ids from: node skills/docentric/scripts/docentric.mjs inspect examples/SalesInvoice.docx)\n` +
    `# node skills/docentric/scripts/docentric.mjs edit examples/SalesInvoice.docx examples/SalesInvoice.edit.yaml --out examples/out/SalesInvoice.edited.docx --preview\n${stringify(ops)}`,
);

await mkdir('examples/ssrs', { recursive: true });
await writeFile('examples/ssrs/ConInvoice.AxReport.xml', AX_REPORT);
await writeFile('examples/ssrs/ConInvoice.spec.yaml', stringify(convertAxReport(AX_REPORT).spec));
console.log('examples/ssrs written');

// README screenshots.
if (process.argv.includes('--images')) {
  await mkdir('docs/assets', { recursive: true });
  const tmp = 'examples/out';
  await mkdir(tmp, { recursive: true });
  const shots: Array<[string, string]> = [];
  const p1 = await renderPreview(created.path, { outDir: tmp, formats: ['png'] });
  if (p1.files.png) shots.push([p1.files.png, 'docs/assets/sales-invoice-preview.png']);
  const p1it = await renderPreview(created.path, { outDir: join(tmp, 'it'), formats: ['png'], language: 'it' });
  if (p1it.files.png) shots.push([p1it.files.png, 'docs/assets/sales-invoice-preview-it.png']);
  const edited = await editTemplate(created.path, ops, { outputPath: join(tmp, 'SalesInvoice.edited.docx') });
  if (edited.path) {
    const p2 = await renderPreview(edited.path, { outDir: tmp, formats: ['png'] });
    if (p2.files.png) shots.push([p2.files.png, 'docs/assets/sales-invoice-edited.png']);
  }
  const ssrs = await createTemplate(convertAxReport(AX_REPORT).spec, join(tmp, 'ConInvoice.docx'));
  const p3 = await renderPreview(ssrs.path, { outDir: tmp, formats: ['png'] });
  if (p3.files.png) shots.push([p3.files.png, 'docs/assets/ssrs-converted.png']);
  for (const [from, to] of shots) {
    await copyFile(from, to);
    console.log(`wrote ${to}`);
  }
  await rm(tmp, { recursive: true, force: true });
}
