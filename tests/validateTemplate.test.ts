import { describe, expect, it } from 'vitest';
import { loadTemplate } from '../src/docx/loadTemplate.js';
import { RULES, renderFindings, validateTemplate } from '../src/validate/validateTemplate.js';
import { buildTemplate, cloneSpec, contosoInvoice, type FixtureSpec, findElement } from './fixtures/builder.js';

async function validate(mutate?: (spec: FixtureSpec) => void) {
  const spec = cloneSpec(contosoInvoice());
  mutate?.(spec);
  return validateTemplate(await loadTemplate(await buildTemplate(spec)));
}

const codes = (r: Awaited<ReturnType<typeof validate>>) => r.findings.map((f) => f.code);

describe('validateTemplate', () => {
  it('passes the clean fixture with no errors or warnings', async () => {
    const r = await validate();
    expect(r.findings.filter((f) => f.severity !== 'info')).toEqual([]);
    expect(r.checkedPaths).toBeGreaterThan(15);
  });

  it('DCX001 content control without element', async () => {
    const r = await validate((s) => {
      s.extraContentControls = [{ tag: '70000001-999', alias: 'Field' }];
    });
    expect(codes(r)).toContain('DCX001');
  });

  it('DCX002 element registered but not placed', async () => {
    const r = await validate((s) => {
      findElement(s, '102').placed = false;
    });
    expect(r.findings.find((f) => f.code === 'DCX002')?.elementId).toBe('102');
  });

  it('DCX003 dangling ChildElements reference', async () => {
    const r = await validate((s) => {
      findElement(s, '110').danglingChildIds = ['555'];
    });
    expect(codes(r)).toContain('DCX003');
  });

  it('DCX004 content control title differs from element type', async () => {
    const r = await validate((s) => {
      findElement(s, '102').alias = 'If';
    });
    expect(codes(r)).toContain('DCX004');
  });

  it('DCX005 unknown data source', async () => {
    const r = await validate((s) => {
      findElement(s, '102').source = 'MainDta';
    });
    expect(codes(r)).toContain('DCX005');
  });

  it("DCX006 Source='.' outside a List", async () => {
    const r = await validate((s) => {
      findElement(s, '102').source = '.';
    });
    expect(codes(r)).toContain('DCX006');
  });

  it('DCX007 missing attribute with a did-you-mean', async () => {
    const r = await validate((s) => {
      findElement(s, '102').path = 'InvoiceHeader/@InvoiceID';
    });
    const f = r.findings.find((x) => x.code === 'DCX007');
    expect(f?.message).toContain("'@InvoiceID' does not exist under MainData:ReportData/InvoiceHeader");
    expect(f?.hint).toContain("'@InvoiceId'");
  });

  it('DCX007 resolves relative paths against the List item, not the root', async () => {
    const r = await validate((s) => {
      findElement(s, '111').path = '@InvoiceId';
    });
    const f = r.findings.find((x) => x.code === 'DCX007');
    expect(f?.message).toContain('ReportData/InvoiceLines');
  });

  it('DCX007 checks the SummaryField aggregate against its collection', async () => {
    const r = await validate((s) => {
      findElement(s, '132').aggregatePath = '@Qty * @Price';
    });
    expect(r.findings.find((x) => x.code === 'DCX007')?.hint).toContain("'@SalesPrice'");
  });

  it('DCX007 missing element step', async () => {
    const r = await validate((s) => {
      findElement(s, '104').path = 'upper-case(Company/@Name)';
    });
    expect(r.findings.find((x) => x.code === 'DCX007')?.hint).toContain("'CurrentCompany'");
  });

  it('DCX008 undefined template label with suggestion', async () => {
    const r = await validate((s) => {
      findElement(s, '131').labelCode = '@DOC_LABEL:Totale';
    });
    const f = r.findings.find((x) => x.code === 'DCX008');
    expect(f?.hint).toContain('@DOC_LABEL:Total');
  });

  it('DCX008 also covers label() inside expressions', async () => {
    const r = await validate((s) => {
      findElement(s, '103').path = 'concat(label("@DOC_LABEL:Client"), InvoiceHeader/@CustName)';
    });
    expect(codes(r)).toContain('DCX008');
  });

  it('DCX009 missing translation', async () => {
    const r = await validate((s) => {
      s.labels[0] = { code: '@DOC_LABEL:InvoiceTitle', texts: { 'en-us': 'Invoice' } };
    });
    expect(r.findings.find((f) => f.code === 'DCX009')?.message).toContain('it');
  });

  it('DCX010 unused labels are reported once', async () => {
    const r = await validate((s) => {
      s.labels.push({ code: '@DOC_LABEL:A', texts: { 'en-us': 'a', it: 'a' } });
      s.labels.push({ code: '@DOC_LABEL:B', texts: { 'en-us': 'b', it: 'b' } });
    });
    const f = r.findings.filter((x) => x.code === 'DCX010');
    expect(f).toHaveLength(1);
    expect(f[0]?.message).toMatch(/^2 of 6/);
  });

  it('ignores cached D365 label texts for translation and usage checks', async () => {
    // The fixture carries @SYS12836 in one language only and never uses it.
    const r = await validate();
    expect(r.findings.some((f) => f.message.includes('@SYS12836'))).toBe(false);
  });

  it('DCX011 D365 labels are listed for the d365fo-mcp server', async () => {
    const r = await validate((s) => {
      findElement(s, '131').labelCode = '@SYS12345';
    });
    expect(r.findings.find((f) => f.code === 'DCX011')?.severity).toBe('info');
  });

  it('DCX012 undeclared variable', async () => {
    const r = await validate((s) => {
      findElement(s, '133').path = "var-element-value('total')";
    });
    expect(r.findings.find((f) => f.code === 'DCX012')?.hint).toContain('tot');
  });

  it('DCX013 element placed twice', async () => {
    const r = await validate((s) => {
      s.duplicatePlacements = ['102'];
    });
    expect(codes(r)).toContain('DCX013');
  });

  it('DCX014 mixed project ids', async () => {
    const r = await validate((s) => {
      s.extraContentControls = [{ tag: '123-102', alias: 'Field' }];
    });
    expect(codes(r)).toContain('DCX014');
  });

  it('DCX015 empty binding', async () => {
    const r = await validate((s) => {
      findElement(s, '102').path = '';
    });
    expect(codes(r)).toContain('DCX015');
  });

  it('DCX016 unparsable expression', async () => {
    const r = await validate((s) => {
      findElement(s, '102').path = 'concat("abc, InvoiceHeader/@InvoiceId)';
    });
    expect(codes(r)).toContain('DCX016');
  });

  it('DCX017 List bound to an expression rather than a collection path', async () => {
    const r = await validate((s) => {
      findElement(s, '121').path = 'iif(@IsWithAmount, InvoiceLines, InvoiceLines)';
    });
    expect(codes(r)).toContain('DCX017');
    // Children of that list now have no item context.
    expect(codes(r)).toContain('DCX006');
  });

  it('sorts errors first and renders a compact report', async () => {
    const r = await validate((s) => {
      findElement(s, '102').path = 'InvoiceHeader/@Nope';
    });
    expect(r.findings[0]?.severity).toBe('error');
    const text = renderFindings(r, 'error');
    expect(text).toMatch(/^Validation: 1 error\(s\)/);
    expect(text).toContain('[error] DCX007');
    expect(text).not.toContain('[info]');
  });

  it('has a rule definition for every code it can emit', () => {
    expect(Object.keys(RULES)).toEqual(Array.from({ length: 19 }, (_, i) => `DCX${String(i + 1).padStart(3, '0')}`));
  });
});
