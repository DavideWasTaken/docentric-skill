import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { parse } from 'yaml';
import { createTemplate } from '../src/create/createTemplate.js';
import { parseLayoutSpec, SpecError } from '../src/create/spec.js';
import { loadTemplateFile } from '../src/docx/loadTemplate.js';
import { WordPackage } from '../src/docx/package.js';
import { fillTemplate } from '../src/preview/fill.js';
import { buildTemplate, contosoInvoice } from './fixtures/builder.js';

let dir: string;
let exampleSpec: unknown;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dcx-create-'));
  exampleSpec = parse(await readFile(resolve('examples/sales-invoice.yaml'), 'utf8'));
  await writeFile(join(dir, 'source.docx'), await buildTemplate(contosoInvoice()));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

describe('createTemplate', () => {
  it('builds the example invoice with no validation errors', async () => {
    const r = await createTemplate(exampleSpec, join(dir, 'invoice.docx'));
    expect(r.validation.counts.error).toBe(0);
    expect(r.provisional).toEqual(['MainData', 'GeneralData']);
    expect(r.validation.findings.map((f) => f.code)).toContain('DCX019');
    expect(r.elements).toBeGreaterThan(25);
  });

  it('registers elements with the Designer hierarchy and colours', async () => {
    const r = await createTemplate(exampleSpec, join(dir, 'invoice2.docx'));
    const t = await loadTemplateFile(r.path);
    const list = [...t.elements.values()].find((e) => e.type === 'List' && e.binding?.path === 'InvoiceLines');
    expect(list?.attributes.RenderMode).toBe('2');
    const children = list?.childIds.map((id) => t.elements.get(id));
    expect(children?.map((c) => c?.binding?.source)).toEqual(['.', '.', '.', '.', '.']);
    // The table sits inside the If on @IsWithAmount.
    const wrapper = [...t.elements.values()].find((e) => e.childIds.includes(list?.id as string));
    expect(wrapper?.type).toBe('If');
    // The SummaryField aggregates the List's collection.
    const sum = [...t.elements.values()].find((e) => e.type === 'SummaryField');
    expect(sum).toMatchObject({
      binding: { source: 'MainData', path: 'InvoiceLines' },
      aggregateBinding: { path: '@Qty * @SalesPrice' },
    });
    // Header content controls exist and the list wraps the table row.
    expect(t.contentControls.some((c) => c.part === 'word/header1.xml')).toBe(true);
    const xml = (await (await WordPackage.openFile(r.path)).text('word/document.xml')) as string;
    expect(xml).toMatch(/<w:alias w:val="List"\/>[\s\S]*?<w15:color w:val="FF00FF"\/><\/w:sdtPr><w:sdtContent><w:tr>/);
  });

  it('produces a document the preview can fill', async () => {
    const r = await createTemplate(exampleSpec, join(dir, 'invoice3.docx'));
    const filled = await fillTemplate(await readFile(r.path));
    expect(filled.stats.listItems).toBe(3);
    expect(filled.issues.filter((i) => !i.startsWith('Images'))).toEqual([]);
  });

  it('copies data sources and cached labels from an existing template', async () => {
    const r = await createTemplate(
      {
        dataSources: { from: 'source.docx' },
        labels: { Hello: { 'en-us': 'Hello', it: 'Ciao' } },
        body: [
          { heading: '{{label:Hello}} {{InvoiceHeader/@CustName}}' },
          { text: '{{label:Customer}}' },
          {
            table: {
              rows: 'InvoiceLines',
              columns: [
                { header: 'Item', value: '{{@ItemId}}' },
                { value: '{{@LineAmount | n2}}', align: 'right' },
              ],
            },
          },
        ],
      },
      join(dir, 'from.docx'),
      { baseDir: dir },
    );
    expect(r.provisional).toEqual([]);
    expect(r.dataSources).toEqual(['GeneralData', 'MainData', 'Parameters']);
    expect(r.validation.counts.error).toBe(0);
    const t = await loadTemplateFile(r.path);
    expect(t.report.dataSourceProviderClass).toBe('ConSalesInvoiceDSP');
    // A label of the source template used by the new layout comes along.
    expect(t.labels.map((l) => l.code)).toEqual(
      expect.arrayContaining(['@DOC_LABEL:Customer', '@DOC_LABEL:Hello', '@SYS12836']),
    );
  });

  it('flags bindings that are not in the schema', async () => {
    const r = await createTemplate(
      { dataSources: { from: 'source.docx' }, body: [{ text: '{{InvoiceHeader/@CustNme}}' }] },
      join(dir, 'typo.docx'),
      { baseDir: dir },
    );
    expect(r.validation.findings.find((f) => f.code === 'DCX007')?.hint).toContain('@CustName');
  });

  it('rejects {{sum:}} outside table totals', async () => {
    await expect(
      createTemplate({ dataSources: { from: 'source.docx' }, body: [{ text: '{{sum: @Qty}}' }] }, join(dir, 'x.docx'), {
        baseDir: dir,
      }),
    ).rejects.toThrow(/totals/);
  });
});

describe('parseLayoutSpec', () => {
  it('names the unknown block', () => {
    expect(() => parseLayoutSpec({ dataSources: { from: 'x' }, body: [{ paragraph: 'x' }] })).toThrow(
      /unknown block \{paragraph\}/,
    );
  });

  it('points at the bad field', () => {
    try {
      parseLayoutSpec({ dataSources: { from: 'x' }, body: [{ table: { rows: 'L', columns: [{ header: 'h' }] } }] });
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(SpecError);
      expect((e as Error).message).toMatch(/body\.0\.table\.columns\.0\.value/);
    }
  });
});
