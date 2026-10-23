import { type Token, tokenize } from './expression.js';

/**
 * Recursive-descent parser for Docentric binding expressions (an XPath 1.0-like
 * language plus Docentric functions). Precedence, lowest first:
 *   or · and · = != · < <= > >= · + - · * div mod idiv · unary - · | · path/primary
 */

export type Axis = 'child' | 'attribute' | 'self' | 'parent' | 'descendant';

export interface PathStep {
  axis: Axis;
  /** Element/attribute name, or '*'. Unused for self/parent. */
  name: string;
  predicates: Expr[];
}

export type Expr =
  | { k: 'str'; v: string }
  | { k: 'num'; v: number }
  | { k: 'var'; name: string }
  | { k: 'call'; name: string; args: Expr[] }
  | { k: 'path'; absolute: boolean; steps: PathStep[] }
  | { k: 'bin'; op: string; l: Expr; r: Expr }
  | { k: 'neg'; e: Expr };

export class ExpressionSyntaxError extends Error {}

const MULTIPLICATIVE_WORDS = new Set(['div', 'mod', 'idiv']);

export function parseExpression(source: string): Expr {
  let tokens: Token[];
  try {
    tokens = tokenize(source);
  } catch (e) {
    throw new ExpressionSyntaxError((e as Error).message);
  }
  let i = 0;
  const peek = (o = 0) => tokens[i + o];
  const isOp = (t: Token | undefined, v: string) => t?.kind === 'op' && t.value === v;
  const isWord = (t: Token | undefined, v: string) => t?.kind === 'name' && t.value === v;
  const fail = (msg: string): never => {
    throw new ExpressionSyntaxError(`${msg} (token ${i + 1} of ${tokens.length})`);
  };
  const expectOp = (v: string) => {
    if (!isOp(peek(), v)) fail(`expected '${v}'`);
    i++;
  };

  const binaryLevel = (next: () => Expr, matches: (t: Token | undefined) => string | null) => (): Expr => {
    let l = next();
    for (let op = matches(peek()); op; op = matches(peek())) {
      i++;
      l = { k: 'bin', op, l, r: next() };
    }
    return l;
  };

  const orExpr = (): Expr => or();
  const unary = (): Expr => {
    if (isOp(peek(), '-')) {
      i++;
      return { k: 'neg', e: unary() };
    }
    return union();
  };
  const multiplicative = binaryLevel(unary, (t) =>
    t?.kind === 'op' && t.value === '*'
      ? '*'
      : t?.kind === 'name' && MULTIPLICATIVE_WORDS.has(t.value)
        ? t.value
        : null,
  );
  const additive = binaryLevel(multiplicative, (t) =>
    t?.kind === 'op' && (t.value === '+' || t.value === '-') ? t.value : null,
  );
  const relational = binaryLevel(additive, (t) =>
    t?.kind === 'op' && ['<', '<=', '>', '>='].includes(t.value) ? t.value : null,
  );
  const equality = binaryLevel(relational, (t) =>
    t?.kind === 'op' && (t.value === '=' || t.value === '!=') ? t.value : null,
  );
  const and = binaryLevel(equality, (t) => (isWord(t, 'and') ? 'and' : null));
  const or = binaryLevel(and, (t) => (isWord(t, 'or') ? 'or' : null));

  function union(): Expr {
    let l = primaryOrPath();
    while (isOp(peek(), '|')) {
      i++;
      l = { k: 'bin', op: '|', l, r: primaryOrPath() };
    }
    return l;
  }

  function primaryOrPath(): Expr {
    const t = peek();
    if (!t) return fail('unexpected end of expression');
    if (t.kind === 'string') {
      i++;
      return { k: 'str', v: t.value };
    }
    if (t.kind === 'number') {
      i++;
      return { k: 'num', v: Number(t.value) };
    }
    if (t.kind === 'var') {
      i++;
      return { k: 'var', name: t.value };
    }
    if (isOp(t, '(')) {
      i++;
      const e = orExpr();
      expectOp(')');
      return e;
    }
    if (t.kind === 'name' && isOp(peek(1), '(')) {
      i += 2;
      const args: Expr[] = [];
      if (!isOp(peek(), ')')) {
        args.push(orExpr());
        while (isOp(peek(), ',')) {
          i++;
          args.push(orExpr());
        }
      }
      expectOp(')');
      return { k: 'call', name: t.value, args };
    }
    return path();
  }

  function path(): Expr {
    const steps: PathStep[] = [];
    let absolute = false;
    let descendantNext = false;
    if (isOp(peek(), '/') || isOp(peek(), '//')) {
      absolute = true;
      descendantNext = (peek() as Token).value === '//';
      i++;
    }
    const first = step(descendantNext);
    if (!first) {
      if (absolute) return { k: 'path', absolute, steps };
      return fail('expected a value, function call or path');
    }
    steps.push(first);
    while (isOp(peek(), '/') || isOp(peek(), '//')) {
      const desc = (peek() as Token).value === '//';
      i++;
      const s = step(desc);
      if (!s) fail('expected a path step');
      steps.push(s as PathStep);
    }
    return { k: 'path', absolute, steps };
  }

  function step(descendant: boolean): PathStep | null {
    const t = peek();
    let s: PathStep | null = null;
    if (isOp(t, '.')) s = { axis: 'self', name: '', predicates: [] };
    else if (isOp(t, '..')) s = { axis: 'parent', name: '', predicates: [] };
    else if (t?.kind === 'attr') s = { axis: 'attribute', name: t.value, predicates: [] };
    else if (t?.kind === 'name' || isOp(t, '*')) {
      s = { axis: descendant ? 'descendant' : 'child', name: (t as Token).value, predicates: [] };
    }
    if (!s) return null;
    i++;
    while (isOp(peek(), '[')) {
      i++;
      s.predicates.push(orExpr());
      expectOp(']');
    }
    return s;
  }

  const result = orExpr();
  if (i < tokens.length) fail(`unexpected '${(peek() as Token).value}'`);
  return result;
}
