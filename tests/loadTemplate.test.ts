import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadTemplate, loadTemplateFile, TemplateError } from '../src/docx/loadTemplate.js';
import { buildTemplate, contosoInvoice } from './fixtures/builder.js';

describe('loadTemplate', () => {
  it('reads report info, languages, labels, data sources and elements', async () => {
    const t = await loadTemplate(await buildTemplate(contosoInvoice()));
    expect(t.productVersion).toBe('6.1.62.0');
    expect(t.report).toMatchObject({ reportId: 'SalesInvoice.Report', dataSourceProviderClass: 'ConSalesInvoiceDSP' });
    expect(t.languages).toEqual(['en-us', 'it']);
    expect(t.labels.find((l) => l.code === '@DOC_LABEL:Total')?.texts).toEqual({ 'en-us': 'Total', it: 'Totale' });
    expect(t.dataSources.map((d) => [d.name, d.section, d.root?.name])).toEqual([
      ['GeneralData', 'GeneralData', 'GeneralData'],
      ['MainData', 'MainData', 'ReportData'],
      ['Parameters', 'Parameters', 'Parameters'],
    ]);
    expect(t.elements.get('110')?.childIds).toEqual(['111', '112', '113', '114']);
    expect(t.elements.get('132')?.aggregateBinding?.path).toBe('@Qty * @SalesPrice');
    expect(t.elements.get('120')?.name).toBe('tot');
  });

  it('reads content controls from body and header with their nesting', async () => {
    const t = await loadTemplate(await buildTemplate(contosoInvoice()));
    const header = t.contentControls.filter((c) => c.part === 'word/header1.xml');
    expect(header.map((c) => c.elementId)).toEqual(['100']);
    expect(t.contentControls[0]?.part).toBe('word/document.xml');
    const field = t.contentControls.find((c) => c.elementId === '116');
    expect(field).toMatchObject({ projectId: '70000001', alias: 'Field', parentElementId: '114' });
    expect(t.contentControls).toHaveLength(t.elements.size);
  });

  it('rejects a Word file without ProjectData', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document xmlns:w="x"/>');
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    await expect(loadTemplate(buf)).rejects.toThrow(/not a Docentric template/);
  });

  it('rejects something that is not a zip', async () => {
    await expect(loadTemplate(Buffer.from('plain text'))).rejects.toBeInstanceOf(TemplateError);
  });

  it('rejects wrong extensions and missing files', async () => {
    await expect(loadTemplateFile('C:/nowhere/template.pdf')).rejects.toThrow(/Not a Word package/);
    await expect(loadTemplateFile('C:/nowhere/template.docx')).rejects.toThrow(/File not found/);
  });
});
