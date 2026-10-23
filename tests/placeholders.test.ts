import { describe, expect, it } from 'vitest';
import { hasPlaceholders, parsePlaceholders } from '../src/write/placeholders.js';

const DS = ['MainData', 'GeneralData', 'Parameters'];

describe('parsePlaceholders', () => {
  it('splits text and fields', () => {
    expect(parsePlaceholders('Order {{Header/@PurchId}}!', DS)).toEqual([
      { kind: 'text', text: 'Order ' },
      { kind: 'field', source: null, expr: 'Header/@PurchId', format: null },
      { kind: 'text', text: '!' },
    ]);
  });

  it('reads an explicit data source only when it exists', () => {
    expect(parsePlaceholders('{{GeneralData: CurrentCompany/@Name}}', DS)[0]).toMatchObject({
      source: 'GeneralData',
      expr: 'CurrentCompany/@Name',
    });
    expect(parsePlaceholders('{{Other: x}}', DS)[0]).toMatchObject({ source: null, expr: 'Other: x' });
  });

  it('reads formats, labels and sums', () => {
    expect(parsePlaceholders('{{@Qty * @Price | n2}}', DS)[0]).toMatchObject({ expr: '@Qty * @Price', format: 'n2' });
    expect(parsePlaceholders('{{Header/@Date | dd/MM/yyyy}}', DS)[0]).toMatchObject({ format: 'dd/MM/yyyy' });
    expect(parsePlaceholders('{{label:Total}}', DS)[0]).toEqual({ kind: 'label', code: '@DOC_LABEL:Total' });
    expect(parsePlaceholders('{{label:@SYS123}}', DS)[0]).toEqual({ kind: 'label', code: '@SYS123' });
    expect(parsePlaceholders('{{sum: @Qty | n2}}', DS)[0]).toEqual({ kind: 'sum', expr: '@Qty', format: 'n2' });
  });

  it('keeps an XPath union that is not a format', () => {
    expect(parsePlaceholders('{{A/@x | B/@y}}', DS)[0]).toMatchObject({ expr: 'A/@x | B/@y', format: null });
  });

  it('detects placeholders', () => {
    expect(hasPlaceholders('a {{b}}')).toBe(true);
    expect(hasPlaceholders('a {b}')).toBe(false);
  });
});
