import { describe, expect, it } from 'vitest';
import { loadTemplate } from '../src/docx/loadTemplate.js';
import { inspectTemplate } from '../src/inspect/inspectTemplate.js';
import { buildTemplate, contosoInvoice } from './fixtures/builder.js';

const load = async () => loadTemplate(await buildTemplate(contosoInvoice()));

describe('inspectTemplate', () => {
  it('summarises report, languages, elements and data sources by default', async () => {
    const text = inspectTemplate(await load());
    expect(text).toContain('Report: SalesInvoice.Report');
    expect(text).toContain('Data source provider class: ConSalesInvoiceDSP');
    expect(text).toContain('Languages: en-us, it');
    expect(text).toContain('- MainData [MainData] root=ReportData');
    expect(text).toContain('## Elements');
    expect(text).not.toContain('## Data source schemas');
  });

  it('nests elements and shows the item context of relative bindings', async () => {
    const text = inspectTemplate(await load(), { include: ['elements'] });
    expect(text).toContain('- List#110 ← MainData: InvoiceLines [RenderMode=2]');
    expect(text).toContain('  - Field#111 ← .(ReportData/InvoiceLines): @ItemId');
    expect(text).toContain('    - Field#116 ← .(ReportData/InvoiceLines): round-number(@DiscPercent, 2)');
    expect(text).toContain('{header1.xml}');
  });

  it('prints the schema tree with a depth limit', async () => {
    const t = await load();
    const shallow = inspectTemplate(t, { include: ['schema'], dataSource: 'maindata', schemaDepth: 0 });
    expect(shallow).toContain('ReportData  @IsWithAmount:boolean');
    expect(shallow).toContain('2 child element(s), raise schemaDepth');
    const deep = inspectTemplate(t, { include: ['schema'], dataSource: 'MainData', schemaDepth: 3 });
    expect(deep).toContain('InvoiceLines[*]  @ItemId:string');
    // Under a <choice maxOccurs="unbounded"> every child is repeatable, as the XSD says.
    expect(deep).toContain('PrintLevel[*]  @Value:integer');
    expect(deep).not.toContain('### GeneralData');
  });

  it('names the available data sources when the filter matches none', async () => {
    const text = inspectTemplate(await load(), { include: ['schema'], dataSource: 'Nope' });
    expect(text).toContain('Available: GeneralData, MainData, Parameters');
  });

  it('lists labels with every language', async () => {
    const text = inspectTemplate(await load(), { include: ['labels'] });
    expect(text).toContain('- @DOC_LABEL:Total [Custom] en-us: "Total" | it: "Totale"');
    expect(text).toContain('Cached D365/ISV label texts (1, resolved again at print time');
    expect(text).not.toContain('- @SYS12836');
  });
});
