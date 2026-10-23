import { DOMParser, type Element } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import {
  type EvalContext,
  evaluateSource,
  formatValue,
  toBoolean,
  toStringValue,
  withLocale,
} from '../src/expr/evaluate.js';
import { ExpressionSyntaxError, parseExpression } from '../src/expr/parser.js';

const SAMPLE = `<ReportData IsWithAmount="true" Flag="false">
  <Header PurchId="PO-001" Notes="" Amount="1234.5000" OrderDate="2026-03-25T00:00:00"><Level Value="2"/></Header>
  <Lines ItemId="A" Qty="2" Price="10.5" Kind="SER"/>
  <Lines ItemId="B" Qty="3" Price="4" Kind="ITM"/>
</ReportData>`;

const root = new DOMParser().parseFromString(SAMPLE, 'text/xml').documentElement as Element;

function ctx(node: Element | null = root, locale = 'en-US'): EvalContext {
  return {
    node,
    root,
    vars: withLocale(new Map(), locale),
    label: (c) => (c === '@DOC_LABEL:Phone' ? 'Tel.' : null),
    unsupported: new Set(),
  };
}

const ev = (expr: string, c = ctx()) => toStringValue(evaluateSource(expr, c));
const lines = () => Array.from(root.getElementsByTagName('Lines'));

describe('parseExpression', () => {
  it('respects precedence', () => {
    expect(parseExpression('1 + 2 * 3')).toMatchObject({ k: 'bin', op: '+', r: { k: 'bin', op: '*' } });
    expect(parseExpression('@a = 1 or @b = 2 and @c = 3')).toMatchObject({ op: 'or', r: { op: 'and' } });
  });

  it('reports syntax errors', () => {
    expect(() => parseExpression('concat(@a,')).toThrow(ExpressionSyntaxError);
    expect(() => parseExpression('@a @b')).toThrow(/unexpected/);
  });
});

describe('evaluate', () => {
  it('reads attributes along paths', () => {
    expect(ev('Header/@PurchId')).toBe('PO-001');
    expect(ev('Header/Level/@Value')).toBe('2');
    expect(ev('@IsWithAmount')).toBe('true');
    expect(ev('/ReportData/Header/@PurchId')).toBe('PO-001');
  });

  it('evaluates relative to the current row', () => {
    const [a] = lines();
    expect(ev('@Qty * @Price', ctx(a as Element))).toBe('21');
    expect(ev('../Header/@PurchId', ctx(a as Element))).toBe('PO-001');
  });

  it('supports predicates and node-set functions', () => {
    expect(ev("Lines[@Kind = 'ITM']/@ItemId")).toBe('B');
    expect(ev('Lines[2]/@ItemId')).toBe('B');
    expect(ev('count(Lines)')).toBe('2');
    expect(ev('sum(Lines/@Qty)')).toBe('5');
  });

  it('runs Docentric functions', () => {
    expect(ev('iif(Header/@Notes = "", "-", Header/@Notes)')).toBe('-');
    expect(ev('concat(Header/@PurchId, $new-line, "x")')).toBe('PO-001\nx');
    expect(ev('upper-case("abc")')).toBe('ABC');
    expect(ev("starts-with(Lines[1]/@Kind, 'SE')")).toBe('true');
    expect(ev('round-number(2.345, 2)')).toBe('2.35');
    expect(ev('replace("a b c", " ", "-")')).toBe('a-b-c');
    expect(ev('label("@DOC_LABEL:Phone")')).toBe('Tel.');
    expect(ev('(10 - 4) div 3')).toBe('2');
  });

  it('reads variables set by Var elements', () => {
    const c = ctx();
    c.vars.set('tot', 42);
    expect(ev("var-element-value('tot') + 1", c)).toBe('43');
  });

  it('converts "false" and "0" to false like the Docentric engine', () => {
    expect(toBoolean(evaluateSource('@Flag', ctx()))).toBe(false);
    expect(toBoolean(evaluateSource('@IsWithAmount', ctx()))).toBe(true);
    expect(toBoolean(evaluateSource('Header/@Missing', ctx()))).toBe(false);
  });

  it('compares node-sets XPath-style', () => {
    expect(ev("Lines/@Kind = 'ITM'")).toBe('true');
    expect(ev('Header/Level/@Value != 2')).toBe('false');
  });

  it('marks unknown functions instead of failing', () => {
    const c = ctx();
    expect(ev('mystery(1)', c)).toBe('⟨mystery()⟩');
    expect([...c.unsupported]).toEqual(['mystery()']);
  });
});

describe('formatValue', () => {
  it('formats numbers with .NET-style patterns in the template culture', () => {
    expect(formatValue(1234.5, 'n2', 'it-IT')).toBe('1.234,50');
    expect(formatValue(1234.5, 'n2', 'en-US')).toBe('1,234.50');
    expect(formatValue(0.256, 'p1', 'en-US')).toBe('25.6 %');
    expect(formatValue('3', '#,##0.00', 'en-US')).toBe('3.00');
  });

  it('formats ISO dates', () => {
    expect(formatValue('2026-03-25T00:00:00', 'dd/MM/yyyy', 'it-IT')).toBe('25/03/2026');
    // The short date pattern comes from ICU and differs between Node builds.
    expect(formatValue('2026-03-25', 'd', 'it-IT')).toMatch(/^25\/0?3\/2026$/);
  });
});
