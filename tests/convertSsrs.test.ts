import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { createTemplate } from '../src/create/createTemplate.js';
import type { Block } from '../src/create/spec.js';
import { fillTemplate } from '../src/preview/fill.js';
import { convertAxReport } from '../src/ssrs/convert.js';
import { AX_REPORT } from './fixtures/axreport.js';

let dir: string;
beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dcx-ssrs-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const find = <K extends string>(blocks: Block[] | undefined, key: K) =>
  (blocks ?? []).find((b) => key in b) as Extract<Block, Record<K, unknown>> | undefined;

describe('convertAxReport', () => {
  const r = convertAxReport(AX_REPORT);

  it('derives a provisional schema from the datasets', () => {
    const schema = (r.spec.dataSources as { schema: Record<string, { elements?: Record<string, unknown> }> }).schema;
    expect(Object.keys(schema)).toEqual(['MainData', 'Parameters', 'GeneralData']);
    expect(schema.MainData?.elements?.ConInvoiceLineTmp).toMatchObject({
      repeated: true,
      attributes: { ItemId: 'string', LineAmount: 'decimal' },
    });
    expect(schema.MainData?.elements?.ConInvoiceHeaderTmp).toMatchObject({
      repeated: false,
      attributes: { InvoiceDate: 'date', CompanyLogo: 'base64Binary' },
    });
  });

  it('turns a Tablix into a data table with header, detail and totals', () => {
    const table = find(r.spec.body, 'table')?.table;
    expect(table?.rows).toBe('ConInvoiceLineTmp');
    expect(table?.columns.map((c) => c.value)).toEqual(['{{@ItemId}}', '{{@Name}}', '{{@LineAmount}}']);
    expect(table?.columns.map((c) => c.header)).toEqual(['Item', '{{label:@SYS7576}}', 'Amount']);
    expect(table?.columns[2]?.align).toBe('right');
    expect(table?.totals?.[0]?.map((c) => (typeof c === 'object' && 'text' in c ? c.text : c))).toEqual([
      'Total',
      '',
      '{{sum: @LineAmount}}',
    ]);
    expect(r.notes.join('\n')).toContain('group row(s) dropped');
  });

  it('translates First(), Format(), concatenation and parameters', () => {
    const grid = find(r.spec.body, 'grid');
    const texts = grid?.grid[0]?.map((c) => (typeof c === 'object' && 'text' in c ? c.text : c));
    expect(texts).toEqual(['{{ConInvoiceHeaderTmp/@InvoiceId}}', '{{ConInvoiceHeaderTmp/@InvoiceDate | d}}']);
    expect(r.spec.body.some((b) => 'text' in b && b.text === 'Bill to: {{ConInvoiceHeaderTmp/@CustName}}')).toBe(true);
    const header = find(r.spec.header, 'grid')?.grid[0]?.map((c) =>
      typeof c === 'object' && 'text' in c ? c.text : c,
    );
    expect(header).toEqual(['{{GeneralData: CurrentCompany/@Name}}', '{{Parameters: @DocumentTitle}}']);
  });

  it('keeps what it cannot translate visible and listed', () => {
    expect(r.spec.body.some((b) => 'text' in b && b.text.startsWith('⟦SSRS: =Code.FormatAddress'))).toBe(true);
    expect(r.notes.some((n) => n.includes('Code.FormatAddress'))).toBe(true);
    expect(r.stats.untranslated).toBe(1);
  });

  it('maps database images to Image elements', () => {
    expect(find(r.spec.body, 'image')).toMatchObject({ image: 'ConInvoiceHeaderTmp/@CompanyLogo', width: 38 });
  });

  it('produces a spec that builds into a valid, previewable template', async () => {
    const created = await createTemplate(r.spec, join(dir, 'converted.docx'));
    expect(created.validation.counts.error).toBe(0);
    const filled = await fillTemplate(await import('node:fs/promises').then((fs) => fs.readFile(created.path)));
    expect(filled.stats.listItems).toBe(3);
  });

  it('rejects files that are not AxReports', () => {
    expect(() => convertAxReport('<AxTable/>')).toThrow(/Not an AxReport/);
  });
});
