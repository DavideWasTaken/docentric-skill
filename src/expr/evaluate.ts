import type { Element, Node } from '@xmldom/xmldom';
import { type Expr, type PathStep, parseExpression } from './parser.js';

/**
 * Evaluates Docentric binding expressions against the sample XML embedded in a
 * template. Used by the preview, so it favours "show something sensible" over
 * strictness: an unknown function yields a visible marker and is reported, it
 * does not abort the document.
 *
 * Values follow XPath 1.0: string | number | boolean | node-set. A node is an
 * element or an attribute of the sample XML.
 */

export type XNode = Element | XAttr;
/** xmldom attributes are Nodes with name/value; typed loosely to stay version-proof. */
export interface XAttr {
  nodeType: 2;
  name: string;
  localName?: string | null;
  value: string;
  ownerElement?: Element | null;
}
export type Value = string | number | boolean | XNode[];

export interface EvalContext {
  /** Current node: the data source root, or the current List item. */
  node: XNode | null;
  /** Root element of the data source the path runs against (for absolute paths). */
  root: Element | null;
  vars: Map<string, Value>;
  label: (code: string) => string | null;
  /** Functions/variables the evaluator met but does not implement. */
  unsupported: Set<string>;
}

const BUILTIN_VARIABLES: Record<string, string> = {
  'new-line': '\n',
  newline: '\n',
  lf: '\n',
  cr: '\n',
  crlf: '\n',
  tab: '\t',
  space: ' ',
  empty: '',
};

const cache = new Map<string, Expr>();

export function evaluateSource(source: string, ctx: EvalContext): Value {
  let ast = cache.get(source);
  if (!ast) {
    ast = parseExpression(source);
    cache.set(source, ast);
  }
  return evaluate(ast, ctx);
}

export function evaluate(e: Expr, ctx: EvalContext): Value {
  switch (e.k) {
    case 'str':
      return e.v;
    case 'num':
      return e.v;
    case 'var': {
      if (e.name in BUILTIN_VARIABLES) return BUILTIN_VARIABLES[e.name] as string;
      const v = ctx.vars.get(e.name);
      if (v !== undefined) return v;
      ctx.unsupported.add(`$${e.name}`);
      return '';
    }
    case 'neg':
      return -toNumber(evaluate(e.e, ctx));
    case 'path':
      return evalPath(e.absolute, e.steps, ctx);
    case 'call':
      return callFunction(e.name, e.args, ctx);
    case 'bin':
      return evalBinary(e.op, e.l, e.r, ctx);
  }
}

function evalBinary(op: string, l: Expr, r: Expr, ctx: EvalContext): Value {
  if (op === 'and') return toBoolean(evaluate(l, ctx)) && toBoolean(evaluate(r, ctx));
  if (op === 'or') return toBoolean(evaluate(l, ctx)) || toBoolean(evaluate(r, ctx));
  const a = evaluate(l, ctx);
  const b = evaluate(r, ctx);
  switch (op) {
    case '|':
      return [...asNodes(a), ...asNodes(b)];
    case '+':
      return toNumber(a) + toNumber(b);
    case '-':
      return toNumber(a) - toNumber(b);
    case '*':
      return toNumber(a) * toNumber(b);
    case 'div':
      return toNumber(a) / toNumber(b);
    case 'idiv':
      return Math.trunc(toNumber(a) / toNumber(b));
    case 'mod':
      return toNumber(a) % toNumber(b);
    default:
      return compare(op, a, b);
  }
}

function compare(op: string, a: Value, b: Value): boolean {
  if (Array.isArray(a)) return a.some((n) => compare(op, stringValue(n), b));
  if (Array.isArray(b)) return b.some((n) => compare(op, a, stringValue(n)));
  if (op === '=' || op === '!=') {
    let eq: boolean;
    if (typeof a === 'boolean' || typeof b === 'boolean') eq = toBoolean(a) === toBoolean(b);
    else if (typeof a === 'number' || typeof b === 'number') eq = toNumber(a) === toNumber(b);
    else eq = a === b;
    return op === '=' ? eq : !eq;
  }
  const x = toNumber(a);
  const y = toNumber(b);
  switch (op) {
    case '<':
      return x < y;
    case '<=':
      return x <= y;
    case '>':
      return x > y;
    default:
      return x >= y;
  }
}

// --- Paths -----------------------------------------------------------------------

function evalPath(absolute: boolean, steps: PathStep[], ctx: EvalContext): XNode[] {
  let current: XNode[];
  let rest = steps;
  if (absolute) {
    // The data root has no document above it here: the first step must name the root itself.
    const first = steps[0];
    if (!first || !ctx.root) return [];
    const matchesRoot = first.name === '*' || first.name === localNameOf(ctx.root);
    const candidates =
      first.axis === 'descendant' ? [ctx.root, ...descendantsOf(ctx.root)] : matchesRoot ? [ctx.root] : [];
    current = applyPredicates(
      candidates.filter((n) => first.name === '*' || localNameOf(n) === first.name),
      first.predicates,
      ctx,
    );
    rest = steps.slice(1);
  } else {
    current = ctx.node ? [ctx.node] : [];
  }
  for (const step of rest) {
    const next: XNode[] = [];
    for (const n of current) next.push(...applyStep(n, step));
    current = applyPredicates(dedupe(next), step.predicates, ctx);
  }
  return current;
}

function applyStep(n: XNode, step: PathStep): XNode[] {
  switch (step.axis) {
    case 'self':
      return [n];
    case 'parent': {
      const p = isAttr(n) ? n.ownerElement : (n.parentNode as Element | null);
      return p && p.nodeType === 1 ? [p] : [];
    }
    case 'attribute': {
      if (isAttr(n)) return [];
      const out: XNode[] = [];
      const attrs = n.attributes;
      for (let i = 0; i < attrs.length; i++) {
        const a = attrs.item(i) as unknown as XAttr | null;
        if (!a || a.name.startsWith('xmlns')) continue;
        if (step.name === '*' || (a.localName ?? a.name) === step.name) out.push(a);
      }
      return out;
    }
    case 'child':
      return isAttr(n) ? [] : childElements(n).filter((c) => step.name === '*' || localNameOf(c) === step.name);
    case 'descendant':
      return isAttr(n) ? [] : descendantsOf(n).filter((c) => step.name === '*' || localNameOf(c) === step.name);
  }
}

function applyPredicates(nodes: XNode[], predicates: Expr[], ctx: EvalContext): XNode[] {
  let out = nodes;
  for (const p of predicates) {
    out = out.filter((n, idx) => {
      const v = evaluate(p, { ...ctx, node: n });
      return typeof v === 'number' ? v === idx + 1 : toBoolean(v);
    });
  }
  return out;
}

function childElements(el: Element): Element[] {
  const out: Element[] = [];
  for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c as Element);
  return out;
}

function descendantsOf(el: Element): Element[] {
  const out: Element[] = [];
  for (const c of childElements(el)) out.push(c, ...descendantsOf(c));
  return out;
}

function dedupe(nodes: XNode[]): XNode[] {
  return [...new Set(nodes)];
}

const isAttr = (n: XNode | Node): n is XAttr => n.nodeType === 2;
const localNameOf = (n: Element) => n.localName ?? n.nodeName;

// --- Conversions -----------------------------------------------------------------

export function stringValue(n: XNode): string {
  return isAttr(n) ? n.value : (n.textContent ?? '');
}

export function toStringValue(v: Value): string {
  if (Array.isArray(v)) return v.length ? stringValue(v[0] as XNode) : '';
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : '';
  return String(v);
}

export function toNumber(v: Value): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const s = toStringValue(v).trim();
  // Sample data leaves amounts empty more often than not; treat that as 0, not NaN.
  if (s === '') return 0;
  const n = Number(s);
  return Number.isNaN(n) ? Number.NaN : n;
}

/**
 * Docentric converts to the binding's result type, so a node whose value is
 * "false" or "0" is false — unlike XPath, where any non-empty node-set is true.
 */
export function toBoolean(v: Value): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0 && !Number.isNaN(v);
  const s = toStringValue(v).trim().toLowerCase();
  return s !== '' && s !== 'false' && s !== '0';
}

export function asNodes(v: Value): XNode[] {
  return Array.isArray(v) ? v : [];
}

// --- Functions -------------------------------------------------------------------

type Fn = (args: Value[], ctx: EvalContext) => Value;

const str = (v: Value | undefined) => (v === undefined ? '' : toStringValue(v));
const num = (v: Value | undefined) => (v === undefined ? 0 : toNumber(v));

const FUNCTIONS: Record<string, Fn> = {
  concat: (a) => a.map(str).join(''),
  iif: (a) => (toBoolean(a[0] ?? false) ? (a[1] ?? '') : (a[2] ?? '')),
  if: (a) => (toBoolean(a[0] ?? false) ? (a[1] ?? '') : (a[2] ?? '')),
  not: (a) => !toBoolean(a[0] ?? false),
  true: () => true,
  false: () => false,
  boolean: (a) => toBoolean(a[0] ?? false),
  number: (a) => num(a[0]),
  string: (a) => str(a[0]),
  'upper-case': (a) => str(a[0]).toUpperCase(),
  'lower-case': (a) => str(a[0]).toLowerCase(),
  'starts-with': (a) => str(a[0]).startsWith(str(a[1])),
  'ends-with': (a) => str(a[0]).endsWith(str(a[1])),
  contains: (a) => str(a[0]).includes(str(a[1])),
  'string-length': (a) => str(a[0]).length,
  'normalize-space': (a) => str(a[0]).replace(/\s+/g, ' ').trim(),
  trim: (a) => str(a[0]).trim(),
  substring: (a) => {
    const s = str(a[0]);
    const start = Math.round(num(a[1])) - 1;
    return a.length > 2 ? s.substr(Math.max(0, start), Math.round(num(a[2]))) : s.slice(Math.max(0, start));
  },
  'substring-before': (a) => {
    const s = str(a[0]);
    const i = s.indexOf(str(a[1]));
    return i < 0 ? '' : s.slice(0, i);
  },
  'substring-after': (a) => {
    const s = str(a[0]);
    const t = str(a[1]);
    const i = s.indexOf(t);
    return i < 0 ? '' : s.slice(i + t.length);
  },
  replace: (a) => {
    const s = str(a[0]);
    const pattern = str(a[1]);
    try {
      return s.replace(new RegExp(pattern, 'g'), str(a[2]));
    } catch {
      return s.split(pattern).join(str(a[2]));
    }
  },
  translate: (a) => {
    const from = str(a[1]);
    const to = str(a[2]);
    return [...str(a[0])]
      .map((c) => {
        const i = from.indexOf(c);
        return i < 0 ? c : (to[i] ?? '');
      })
      .join('');
  },
  char: (a) => String.fromCharCode(num(a[0])),
  'string-join': (a) =>
    asNodes(a[0] ?? [])
      .map(stringValue)
      .join(str(a[1])),
  'round-number': (a) => roundTo(num(a[0]), a.length > 1 ? num(a[1]) : 0),
  round: (a) => roundTo(num(a[0]), a.length > 1 ? num(a[1]) : 0),
  floor: (a) => Math.floor(num(a[0])),
  ceiling: (a) => Math.ceil(num(a[0])),
  abs: (a) => Math.abs(num(a[0])),
  sum: (a) => asNodes(a[0] ?? []).reduce((t, n) => t + toNumber([n]), 0),
  count: (a) => asNodes(a[0] ?? []).length,
  avg: (a) => {
    const n = asNodes(a[0] ?? []);
    return n.length ? n.reduce((t, x) => t + toNumber([x]), 0) / n.length : 0;
  },
  min: (a) => Math.min(...asNodes(a[0] ?? []).map((n) => toNumber([n]))),
  max: (a) => Math.max(...asNodes(a[0] ?? []).map((n) => toNumber([n]))),
  exists: (a) => asNodes(a[0] ?? []).length > 0,
  empty: (a) => asNodes(a[0] ?? []).length === 0,
  format: (a) => formatValue(a[0] ?? '', str(a[1]), 'en-US'),
  'format-number': (a) => formatValue(a[0] ?? '', str(a[1]), 'en-US'),
  label: (a, ctx) => {
    const code = str(a[0]);
    return ctx.label(code) ?? code;
  },
  'var-element-value': (a, ctx) => ctx.vars.get(str(a[0])) ?? '',
};

function callFunction(name: string, argExprs: Expr[], ctx: EvalContext): Value {
  const fn = FUNCTIONS[name];
  if (!fn) {
    ctx.unsupported.add(`${name}()`);
    return `⟨${name}()⟩`;
  }
  const args = argExprs.map((a) => evaluate(a, ctx));
  // `format` follows the template culture, which the preview injects through the label hook.
  if (name === 'format' || name === 'format-number') {
    return formatValue(args[0] ?? '', str(args[1]), currentLocale(ctx));
  }
  return fn(args, ctx);
}

function roundTo(n: number, digits: number): number {
  const f = 10 ** digits;
  return Math.round(n * f) / f;
}

// --- Formatting ------------------------------------------------------------------

const LOCALE_KEY = '\u0000locale';

/** The preview stores the template culture in the variable map under a reserved key. */
export function withLocale(vars: Map<string, Value>, locale: string): Map<string, Value> {
  vars.set(LOCALE_KEY, locale);
  return vars;
}

function currentLocale(ctx: EvalContext): string {
  const v = ctx.vars.get(LOCALE_KEY);
  return typeof v === 'string' ? v : 'en-US';
}

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?/;

/**
 * .NET-style format strings, the subset Docentric templates use: n0–n9, f0–f9,
 * p0–p9, d/D for dates, and custom patterns (#,##0.00 · dd/MM/yyyy).
 */
export function formatValue(v: Value, format: string, locale: string): string {
  const s = toStringValue(v);
  const date = ISO_DATE.exec(s);
  if (date) return formatDate(date, format, locale);
  if (!format) return typeof v === 'number' ? formatPlain(v, locale) : s;
  const n = toNumber(v);
  if (Number.isNaN(n)) return s;
  const m = /^([nNfFpPcC])(\d*)$/.exec(format);
  if (m) {
    const kind = (m[1] as string).toLowerCase();
    const digits = m[2] ? Number(m[2]) : 2;
    const opts: Intl.NumberFormatOptions = {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
      useGrouping: kind === 'n' || kind === 'c',
    };
    if (kind === 'p') return `${new Intl.NumberFormat(locale, opts).format(n * 100)} %`;
    return new Intl.NumberFormat(locale, opts).format(n);
  }
  if (/^[#0,.]+$/.test(format)) {
    const decimals = format.includes('.') ? (format.split('.')[1] ?? '').length : 0;
    const minDecimals = format.includes('.') ? (format.split('.')[1] ?? '').replace(/#/g, '').length : 0;
    return new Intl.NumberFormat(locale, {
      minimumFractionDigits: minDecimals,
      maximumFractionDigits: decimals,
      useGrouping: format.includes(','),
    }).format(n);
  }
  return s;
}

/** Plain number display (.NET ToString()): culture decimal separator, no grouping, no trailing zeros. */
export function formatPlain(n: number, locale: string): string {
  if (!Number.isFinite(n)) return '';
  return new Intl.NumberFormat(locale, { maximumFractionDigits: 10, useGrouping: false }).format(n);
}

function formatDate(m: RegExpExecArray, format: string, locale: string): string {
  const [y, mo, d, h = '00', mi = '00', se = '00'] = m.slice(1).map((x) => x ?? '00') as string[];
  const date = new Date(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), Number(se));
  if (!format || format === 'd') return date.toLocaleDateString(locale);
  if (format === 'D') return date.toLocaleDateString(locale, { dateStyle: 'long' });
  if (format === 'g' || format === 'G') return date.toLocaleString(locale);
  return format
    .replace(/yyyy/g, y as string)
    .replace(/yy/g, (y as string).slice(2))
    .replace(/MM/g, mo as string)
    .replace(/dd/g, d as string)
    .replace(/HH/g, h)
    .replace(/mm/g, mi)
    .replace(/ss/g, se);
}
