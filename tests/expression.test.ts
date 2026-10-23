import { describe, expect, it } from 'vitest';
import { analyzeExpression, tokenize } from '../src/expr/expression.js';

const pathTexts = (expr: string) => analyzeExpression(expr).paths.map((p) => p.text);

describe('tokenize', () => {
  it('keeps hyphenated function names whole and splits spaced minus', () => {
    const kinds = tokenize("var-element-value('x') - 1").map((t) => `${t.kind}:${t.value}`);
    expect(kinds).toEqual(['name:var-element-value', 'op:(', 'string:x', 'op:)', 'op:-', 'number:1']);
  });

  it('reads attributes, variables and two-char operators', () => {
    const kinds = tokenize('@Qty != $new-line').map((t) => t.kind);
    expect(kinds).toEqual(['attr', 'op', 'var']);
  });

  it('rejects an unterminated string', () => {
    expect(() => tokenize('concat("abc, @X)')).toThrow(/Unterminated string/);
  });
});

describe('analyzeExpression', () => {
  it('extracts paths from a plain path', () => {
    expect(pathTexts('InvoiceHeader/@InvoiceId')).toEqual(['InvoiceHeader/@InvoiceId']);
  });

  it('extracts paths from function arguments, not function names', () => {
    const info = analyzeExpression('iif(InvoiceHeader/@Mode = "", "-", InvoiceHeader/@Mode)');
    expect(info.paths.map((p) => p.text)).toEqual(['InvoiceHeader/@Mode', 'InvoiceHeader/@Mode']);
    expect(info.functions).toEqual(['iif']);
  });

  it('treats and/or/div as operators after an operand', () => {
    expect(pathTexts('(@A - Header/@B) div @A * 100')).toEqual(['@A', 'Header/@B', '@A']);
    expect(pathTexts('@X != "" or Header/Level/@Value = 2')).toEqual(['@X', 'Header/Level/@Value']);
  });

  it('treats an element literally named like an operator as a path at operand position', () => {
    expect(pathTexts('or/@x')).toEqual(['or/@x']);
  });

  it('collects label codes and variable references', () => {
    const info = analyzeExpression(
      `concat(label("@DOC_LABEL:Phone"), ": ", round-number(var-element-value('tot'), 2), $lf)`,
    );
    expect(info.labelCodes).toEqual(['@DOC_LABEL:Phone']);
    expect(info.varRefs).toEqual(['tot']);
    expect(info.variables).toEqual(['lf']);
    expect(info.paths).toEqual([]);
  });

  it('marks predicates and descendant paths', () => {
    const [pred] = analyzeExpression('Lines[@Qty > 0]/@ItemId').paths;
    expect(pred?.predicate).toBe(true);
    expect(pred?.steps.map((s) => s.kind)).toEqual(['element', 'attribute']);
    const [desc] = analyzeExpression('//Lines/@ItemId').paths;
    expect(desc?.descendant).toBe(true);
    expect(desc?.absolute).toBe(true);
  });

  it('distinguishes multiplication from a wildcard step', () => {
    expect(pathTexts('@Qty * @Price')).toEqual(['@Qty', '@Price']);
    const [wild] = analyzeExpression('Header/*').paths;
    expect(wild?.steps.at(-1)?.kind).toBe('wildcard');
  });

  it('reads relative parent steps', () => {
    const [p] = analyzeExpression('../@CurrencyCode').paths;
    expect(p?.steps.map((s) => s.kind)).toEqual(['parent', 'attribute']);
  });

  it('reports a parse error instead of throwing', () => {
    expect(analyzeExpression('concat("x').error).toMatch(/Unterminated/);
  });
});
