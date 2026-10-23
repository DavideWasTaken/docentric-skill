/**
 * Static reading of Docentric binding expressions — an XPath-like language with
 * Docentric functions (iif, concat, format, label, var-element-value, ...).
 *
 * This is not an evaluator. It extracts what can be checked offline:
 *   - location paths (PurchOrderHeader/@PurchId, @Qty, ../Lines) → resolved against the XSD
 *   - label codes passed to label("...")
 *   - variable names passed to var-element-value('...')
 *   - function names and $variables, for inspection
 * Predicates ([...]) are skipped: the path before them is still checked, what is
 * inside is not.
 */

export type Token =
  | { kind: 'string'; value: string }
  | { kind: 'number'; value: string }
  | { kind: 'name'; value: string }
  | { kind: 'attr'; value: string }
  | { kind: 'var'; value: string }
  | { kind: 'op'; value: string };

export type Step =
  | { kind: 'element'; name: string }
  | { kind: 'attribute'; name: string }
  | { kind: 'self' }
  | { kind: 'parent' }
  | { kind: 'wildcard' };

export interface PathRef {
  text: string;
  absolute: boolean;
  steps: Step[];
  /** Contains `//`: the steps cannot be resolved one level at a time. */
  descendant: boolean;
  /** Had one or more [predicates]; the steps are still checked, the predicates are not. */
  predicate: boolean;
}

export interface ExpressionInfo {
  paths: PathRef[];
  functions: string[];
  labelCodes: string[];
  varRefs: string[];
  variables: string[];
  error: string | null;
}

const NAME_START = /[A-Za-z_]/;
const NAME_CHAR = /[A-Za-z0-9_.-]/;
const OPERATOR_WORDS = new Set(['and', 'or', 'div', 'mod', 'idiv', 'eq', 'ne', 'lt', 'le', 'gt', 'ge']);

export function tokenize(expr: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const readName = () => {
    const start = i;
    while (i < expr.length && NAME_CHAR.test(expr[i] as string)) i++;
    // A trailing '-' or '.' belongs to the next token (e.g. `a- b`).
    while (i > start + 1 && /[-.]/.test(expr[i - 1] as string)) i--;
    return expr.slice(start, i);
  };
  while (i < expr.length) {
    const c = expr[i] as string;
    if (/\s/.test(c)) {
      i++;
    } else if (c === '"' || c === "'") {
      const end = expr.indexOf(c, i + 1);
      if (end < 0) throw new Error(`Unterminated string starting at ${i}`);
      tokens.push({ kind: 'string', value: expr.slice(i + 1, end) });
      i = end + 1;
    } else if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(expr[i + 1] ?? ''))) {
      const m = /^[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?/.exec(expr.slice(i));
      const value = m?.[0] ?? c;
      tokens.push({ kind: 'number', value });
      i += value.length;
    } else if (c === '@') {
      i++;
      if (expr[i] === '*') {
        tokens.push({ kind: 'attr', value: '*' });
        i++;
      } else {
        tokens.push({ kind: 'attr', value: readName() });
      }
    } else if (c === '$') {
      i++;
      tokens.push({ kind: 'var', value: readName() });
    } else if (NAME_START.test(c)) {
      tokens.push({ kind: 'name', value: readName() });
    } else {
      const two = expr.slice(i, i + 2);
      if (['!=', '<=', '>=', '//', '..'].includes(two)) {
        tokens.push({ kind: 'op', value: two });
        i += 2;
      } else if ('()[],/=<>+-*|.'.includes(c)) {
        tokens.push({ kind: 'op', value: c });
        i++;
      } else {
        throw new Error(`Unexpected character '${c}' at ${i}`);
      }
    }
  }
  return tokens;
}

const isOp = (t: Token | undefined, value: string) => t?.kind === 'op' && t.value === value;

/** True when the token closes an operand, so what follows is in operator position. */
function endsOperand(t: Token | undefined): boolean {
  if (!t) return false;
  if (t.kind === 'op') return t.value === ')' || t.value === ']' || t.value === '.' || t.value === '..';
  return true;
}

export function analyzeExpression(expr: string): ExpressionInfo {
  const info: ExpressionInfo = { paths: [], functions: [], labelCodes: [], varRefs: [], variables: [], error: null };
  let tokens: Token[];
  try {
    tokens = tokenize(expr);
  } catch (e) {
    info.error = (e as Error).message;
    return info;
  }

  let i = 0;
  const skipPredicate = () => {
    // tokens[i] is '['; skip to the matching ']'
    let depth = 0;
    for (; i < tokens.length; i++) {
      if (isOp(tokens[i], '[')) depth++;
      else if (isOp(tokens[i], ']') && --depth === 0) {
        i++;
        return;
      }
    }
  };

  const readStep = (): Step | null => {
    const t = tokens[i];
    if (!t) return null;
    if (t.kind === 'name' && !isOp(tokens[i + 1], '(')) {
      i++;
      return { kind: 'element', name: t.value };
    }
    if (t.kind === 'attr') {
      i++;
      return t.value === '*' ? { kind: 'wildcard' } : { kind: 'attribute', name: t.value };
    }
    if (isOp(t, '.')) {
      i++;
      return { kind: 'self' };
    }
    if (isOp(t, '..')) {
      i++;
      return { kind: 'parent' };
    }
    if (isOp(t, '*')) {
      i++;
      return { kind: 'wildcard' };
    }
    return null;
  };

  while (i < tokens.length) {
    const t = tokens[i] as Token;
    const prev = tokens[i - 1];

    if (t.kind === 'name' && isOp(tokens[i + 1], '(')) {
      info.functions.push(t.value);
      const arg = tokens[i + 2];
      if (arg?.kind === 'string') {
        if (t.value === 'label') info.labelCodes.push(arg.value);
        if (t.value === 'var-element-value') info.varRefs.push(arg.value);
      }
      i += 2;
      continue;
    }
    if (t.kind === 'name' && OPERATOR_WORDS.has(t.value) && endsOperand(prev)) {
      i++;
      continue;
    }
    if (t.kind === 'var') {
      info.variables.push(t.value);
      i++;
      continue;
    }

    const startsPath =
      t.kind === 'name' ||
      t.kind === 'attr' ||
      isOp(t, '.') ||
      isOp(t, '..') ||
      ((isOp(t, '/') || isOp(t, '//')) && !endsOperand(prev));
    if (!startsPath) {
      i++;
      continue;
    }

    const start = i;
    const ref: PathRef = { text: '', absolute: false, steps: [], descendant: false, predicate: false };
    if (isOp(t, '/') || isOp(t, '//')) {
      ref.absolute = true;
      ref.descendant = t.value === '//';
      i++;
    }
    const first = readStep();
    if (first) ref.steps.push(first);
    while (i < tokens.length) {
      if (isOp(tokens[i], '[')) {
        ref.predicate = true;
        skipPredicate();
        continue;
      }
      if (isOp(tokens[i], '/') || isOp(tokens[i], '//')) {
        if ((tokens[i] as Token).value === '//') ref.descendant = true;
        i++;
        const step = readStep();
        if (step) ref.steps.push(step);
        continue;
      }
      break;
    }
    if (i === start) i++;
    ref.text = render(tokens.slice(start, i));
    if (ref.steps.length > 0 || ref.absolute) info.paths.push(ref);
  }

  info.functions = [...new Set(info.functions)];
  info.variables = [...new Set(info.variables)];
  return info;
}

function render(tokens: Token[]): string {
  return tokens.map((t) => (t.kind === 'attr' ? `@${t.value}` : t.kind === 'var' ? `$${t.value}` : t.value)).join('');
}
