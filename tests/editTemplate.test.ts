import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { parse } from 'yaml';
import { createTemplate } from '../src/create/createTemplate.js';
import { loadTemplateFile } from '../src/docx/loadTemplate.js';
import { editTemplate, resizeSpan, takeSpace } from '../src/edit/editTemplate.js';
import { fillTemplate } from '../src/preview/fill.js';
import { buildTemplate, contosoInvoice } from './fixtures/builder.js';

let dir: string;
let fixture: string;
let invoice: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dcx-edit-'));
  fixture = join(dir, 'fixture.docx');
  await writeFile(fixture, await buildTemplate(contosoInvoice()));
  invoice = (
    await createTemplate(parse(await readFile('examples/sales-invoice.yaml', 'utf8')), join(dir, 'invoice.docx'))
  ).path;
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const out = (name: string) => join(dir, name);

async function listOf(path: string) {
  const t = await loadTemplateFile(path);
  return [...t.elements.values()].find((e) => e.type === 'List' && e.attributes.RenderMode === '2');
}

describe('editTemplate', () => {
  it('adds and updates labels', async () => {
    const r = await editTemplate(
      fixture,
      [
        { op: 'set_label', code: 'Unit', texts: { 'en-us': 'Unit', it: 'Unità' } },
        { op: 'set_label', code: '@DOC_LABEL:Total', texts: { it: 'Totale generale' } },
      ],
      { outputPath: out('labels.docx') },
    );
    expect(r.written).toBe(true);
    const t = await loadTemplateFile(out('labels.docx'));
    expect(t.labels.find((l) => l.code === '@DOC_LABEL:Unit')?.texts).toEqual({ 'en-us': 'Unit', it: 'Unità' });
    expect(t.labels.find((l) => l.code === '@DOC_LABEL:Total')?.texts.it).toBe('Totale generale');
  });

  it('takes the space for a new column from the wide columns, keeping the table width', () => {
    // Item 2000, Description 5000, Qty 1000, Price 1000 → add 1000 after Item.
    const out = takeSpace([2000, 5000, 1000, 1000], 1, 1000);
    expect(out.reduce((a, b) => a + b, 0)).toBe(9000);
    expect(out[1]).toBe(1000);
    // Narrow numeric columns keep their width; Description gives way.
    expect(out.slice(3)).toEqual([1000, 1000]);
    expect(out[2]).toBeLessThan(5000);
    // When the wide columns cannot give enough, everything shrinks proportionally.
    const tight = takeSpace([1000, 1000, 1000], 0, 1500);
    expect(tight.reduce((a, b) => a + b, 0)).toBe(3000);
  });

  it('resizes a column span, keeping the table width', () => {
    // Grow column 0 from 1000 to 2000: the 1000 comes from the wide column (5000), not the narrow ones.
    const grown = resizeSpan([1000, 5000, 1000, 1000], 0, 1, 2000);
    expect(grown).toEqual([2000, 4000, 1000, 1000]);
    // Shrink it back: the freed space goes to the widest other column.
    expect(resizeSpan([2000, 4000, 1000, 1000], 0, 1, 1000)).toEqual([1000, 5000, 1000, 1000]);
    // A spanning cell is resized as a whole, in proportion.
    const spanned = resizeSpan([1000, 1000, 6000], 0, 2, 3000);
    expect(spanned.slice(0, 2)).toEqual([1500, 1500]);
    expect(spanned.reduce((a, b) => a + b, 0)).toBe(8000);
  });

  it('resize_column sets a column width in cm and reports what gave way', async () => {
    const list = await listOf(invoice);
    const itemField = list?.childIds[0];
    const r = await editTemplate(invoice, [{ op: 'resize_column', element: `#${itemField}`, widthCm: 4 }], {
      outputPath: out('resized.docx'),
    });
    expect(r.written).toBe(true);
    expect(r.applied[0]?.detail).toMatch(/→ 4 cm; other columns: .*cm/);
    const xml = (await (
      await JSZip.loadAsync(await readFile(out('resized.docx')))
    )
      .file('word/document.xml')
      ?.async('string')) as string;
    expect(xml).toContain('<w:gridCol w:w="2268"/>');
  });

  it('adds a table column with header and value, keeping the grid consistent', async () => {
    const list = await listOf(invoice);
    const itemField = list?.childIds[0];
    const r = await editTemplate(
      invoice,
      [
        {
          op: 'add_column',
          list: `#${list?.id}`,
          after: `#${itemField}`,
          header: 'Line amount',
          value: '{{@LineAmount}}',
          format: 'n2',
        },
      ],
      { outputPath: out('column.docx') },
    );
    expect(r.newErrors).toEqual([]);
    expect(r.applied[0]?.newElements.length).toBe(1);
    const after = await listOf(out('column.docx'));
    expect(after?.childIds).toHaveLength(6);
    const xml = (await (
      await JSZip.loadAsync(await readFile(out('column.docx')))
    )
      .file('word/document.xml')
      ?.async('string')) as string;
    const table = xml.slice(xml.indexOf('Line amount') - 4000, xml.indexOf('Line amount') + 4000);
    const grid =
      /<w:tblGrid>(.*?)<\/w:tblGrid>/.exec(xml.slice(xml.lastIndexOf('<w:tbl>', xml.indexOf('Line amount'))))?.[1] ??
      '';
    expect((grid.match(/<w:gridCol /g) ?? []).length).toBe(6);
    expect(table).toContain('@LineAmount');
    // The preview shows the new column for every row.
    const filled = await fillTemplate(await readFile(out('column.docx')));
    const text = (await filled.pkg.text('word/document.xml')) as string;
    expect(text).toContain('Line amount');
  });

  it('inserts blocks after an element and registers them under its parent', async () => {
    const r = await editTemplate(
      fixture,
      [{ op: 'insert', anchor: { element: '#111' }, blocks: [{ text: 'Name: {{@Name}}' }] }],
      { outputPath: out('insert.docx') },
    );
    const id = r.applied[0]?.newElements[0] as string;
    const t = await loadTemplateFile(out('insert.docx'));
    expect(t.elements.get(id)?.binding).toMatchObject({ source: '.', path: '@Name' });
    expect(t.elements.get('110')?.childIds).toContain(id);
    expect(r.after.counts.error).toBe(0);
  });

  it('inserts next to a paragraph found by its design-time text', async () => {
    // In a template, a field shows its binding path as text: that is what an anchor matches.
    const r = await editTemplate(
      invoice,
      [
        {
          op: 'insert',
          anchor: { text: 'InvoiceHeader/@CustAddress' },
          position: 'before',
          blocks: [{ text: 'VAT {{GeneralData: CurrentCompany/@VATNum}}' }],
        },
      ],
      { outputPath: out('by-text.docx') },
    );
    expect(r.written).toBe(true);
    const t = await loadTemplateFile(out('by-text.docx'));
    expect(
      [...t.elements.values()].some(
        (e) => e.binding?.source === 'GeneralData' && e.binding.path === 'CurrentCompany/@VATNum',
      ),
    ).toBe(true);
  });

  it('wraps an element in an If and moves it under the new If', async () => {
    const r = await editTemplate(
      fixture,
      [{ op: 'wrap_if', element: '#102', condition: 'InvoiceHeader/@InvoiceId != ""' }],
      {
        outputPath: out('wrap.docx'),
      },
    );
    const ifId = r.applied[0]?.newElements[0] as string;
    const t = await loadTemplateFile(out('wrap.docx'));
    expect(t.elements.get(ifId)).toMatchObject({ type: 'If', childIds: ['102'] });
    expect(t.contentControls.find((c) => c.elementId === '102')?.parentElementId).toBe(ifId);
  });

  it('removes an element with its children', async () => {
    const r = await editTemplate(fixture, [{ op: 'remove', element: '#114' }], { outputPath: out('remove.docx') });
    const t = await loadTemplateFile(out('remove.docx'));
    expect(t.elements.has('114')).toBe(false);
    expect(t.elements.has('116')).toBe(false);
    expect(t.elements.get('110')?.childIds).toEqual(['111', '112', '113']);
    expect(r.after.counts.error).toBe(0);
  });

  it('changes a binding and its display text', async () => {
    await editTemplate(
      fixture,
      [{ op: 'set_binding', element: '#102', path: 'InvoiceHeader/@CustName', format: 'n0' }],
      {
        outputPath: out('binding.docx'),
      },
    );
    const t = await loadTemplateFile(out('binding.docx'));
    expect(t.elements.get('102')).toMatchObject({
      binding: { path: 'InvoiceHeader/@CustName' },
      attributes: { FormatString: 'n0' },
    });
  });

  it('converts {{placeholders}} typed in Word into elements', async () => {
    const zip = await JSZip.loadAsync(await readFile(fixture));
    const xml = (await zip.file('word/document.xml')?.async('string')) as string;
    zip.file(
      'word/document.xml',
      xml.replace(
        '<w:body>',
        '<w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Dear {{InvoiceHeader/@CustNa</w:t></w:r><w:r><w:t>me}}, {{label:Total}}</w:t></w:r></w:p>',
      ),
    );
    const typed = out('typed.docx');
    await writeFile(typed, await zip.generateAsync({ type: 'nodebuffer' }));
    const r = await editTemplate(typed, [{ op: 'convert_placeholders' }], { outputPath: out('converted.docx') });
    expect(r.applied[0]?.detail).toMatch(/^2 placeholder/);
    const t = await loadTemplateFile(out('converted.docx'));
    const created = r.applied[0]?.newElements.map((id) => t.elements.get(id));
    expect(created?.map((e) => e?.type)).toEqual(['Field', 'Label']);
    expect(created?.[0]?.binding).toMatchObject({ source: 'MainData', path: 'InvoiceHeader/@CustName' });
    const doc = (await (
      await JSZip.loadAsync(await readFile(out('converted.docx')))
    )
      .file('word/document.xml')
      ?.async('string')) as string;
    expect(doc).toMatch(/<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Dear <\/w:t><\/w:r>/);
  });

  it('refuses to save an edit that introduces errors, unless forced', async () => {
    const bad = [{ op: 'set_binding', element: '#102', path: 'InvoiceHeader/@Nope' }];
    const r = await editTemplate(fixture, bad, { outputPath: out('bad.docx') });
    expect(r.written).toBe(false);
    expect(r.newErrors[0]?.code).toBe('DCX007');
    const forced = await editTemplate(fixture, bad, { outputPath: out('bad.docx'), force: true });
    expect(forced.written).toBe(true);
  });

  it('never overwrites the template and explains bad operations', async () => {
    await expect(editTemplate(fixture, [{ op: 'remove', element: '#102' }], { outputPath: fixture })).rejects.toThrow(
      /must differ/,
    );
    await expect(editTemplate(fixture, [{ op: 'remove', element: 'x' }])).rejects.toThrow(/element id/);
    await expect(
      editTemplate(fixture, [{ op: 'remove', element: '#999' }], { outputPath: out('nope.docx') }),
    ).rejects.toThrow(/no element 999/);
  });

  it('adds a language, and validation then asks for its texts', async () => {
    const r = await editTemplate(fixture, [{ op: 'add_language', language: 'DE', default: true }], {
      outputPath: out('de.docx'),
    });
    const t = await loadTemplateFile(out('de.docx'));
    expect(t.languages).toEqual(['en-us', 'it', 'de']);
    expect(t.defaultLanguage).toBe('de');
    expect(r.after.findings.filter((f) => f.code === 'DCX009').length).toBeGreaterThan(0);
  });

  it('supports dry runs', async () => {
    const r = await editTemplate(fixture, [{ op: 'remove', element: '#102' }], {
      outputPath: out('dry.docx'),
      dryRun: true,
    });
    expect(r.written).toBe(false);
    expect(r.applied).toHaveLength(1);
  });
});
