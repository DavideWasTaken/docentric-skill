import { describe, expect, it } from 'vitest';
import { fillTemplate, toLocale } from '../src/preview/fill.js';
import { renderHtml } from '../src/preview/html.js';
import { buildTemplate, cloneSpec, contosoInvoice, type FixtureDataSource, findElement } from './fixtures/builder.js';

async function filled(
  mutate?: Parameters<typeof cloneSpec>[0] extends infer S ? (s: S) => void : never,
  language?: string,
) {
  const spec = cloneSpec(contosoInvoice());
  mutate?.(spec);
  return fillTemplate(await buildTemplate(spec), { language });
}

async function bodyText(r: Awaited<ReturnType<typeof filled>>): Promise<string> {
  const xml = (await r.pkg.text('word/document.xml')) as string;
  return [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('|');
}

describe('fillTemplate', () => {
  it('replaces fields and labels with sample values and removes the content controls', async () => {
    const r = await filled();
    const text = await bodyText(r);
    expect(text).toContain('INV-0001');
    expect(text).toContain('Customer: Fabrikam');
    expect(text).toContain('CONTOSO LTD');
    expect(text).toContain('Invoice');
    expect(await r.pkg.text('word/document.xml')).not.toContain('<w:sdt>');
  });

  it('repeats list rows, computes variables and totals', async () => {
    const r = await filled((s) => {
      (s.dataSources[1] as FixtureDataSource).sample =
        '<ReportData IsWithAmount="true"><InvoiceHeader InvoiceId="X" CustName="F" CurrencyCode="EUR"><PrintLevel Value="1"/></InvoiceHeader>' +
        '<InvoiceLines ItemId="A" Name="One" Qty="2" SalesPrice="10" LineAmount="20" DiscPercent="0"/>' +
        '<InvoiceLines ItemId="B" Name="Two" Qty="1" SalesPrice="5" LineAmount="5" DiscPercent="10"/></ReportData>';
    });
    const text = await bodyText(r);
    expect(r.stats.listItems).toBe(4); // two lists × two lines
    expect(text).toContain('A|');
    expect(text).toContain('B|');
    expect(text).toContain('25'); // SummaryField Σ Qty*SalesPrice and Var tot = 25
    expect(text).toContain('Discount'); // the If on DiscPercent > 0 holds for line B only
    expect(r.stats.ifsTrue).toBeGreaterThanOrEqual(1);
  });

  it('drops content whose condition is false', async () => {
    const r = await filled((s) => {
      findElement(s, '130').path = '@IsWithAmount = "nope"';
    });
    expect(await bodyText(r)).not.toContain('Total');
  });

  it('uses the requested language for labels', async () => {
    const r = await filled(undefined, 'it');
    expect(await bodyText(r)).toContain('Fattura');
    expect(r.locale).toBe('it-IT');
  });

  it('removes the Docentric part so the preview is a plain document', async () => {
    const r = await filled();
    expect(r.pkg.has('customXml/item1.xml')).toBe(false);
    expect(await r.pkg.text('word/_rels/document.xml.rels')).not.toContain('customXml/item1.xml');
  });

  it('reports expressions it cannot evaluate', async () => {
    const r = await filled((s) => {
      findElement(s, '102').path = 'mystery(InvoiceHeader/@InvoiceId)';
    });
    expect(r.issues.join('\n')).toContain('mystery()');
  });

  it('renders an HTML preview with the header part', async () => {
    const r = await filled();
    const html = await renderHtml(r.pkg, 'test');
    expect(html).toContain('class="hdr"');
    expect(html).toContain('INV-0001');
  });
});

describe('toLocale', () => {
  it('maps template languages to locales', () => {
    expect(toLocale('it')).toBe('it-IT');
    expect(toLocale('en-us')).toBe('en-US');
    expect(toLocale('de-at')).toBe('de-AT');
  });
});
