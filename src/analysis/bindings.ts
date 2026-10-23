import { analyzeExpression, type ExpressionInfo, type PathRef } from '../expr/expression.js';
import type { Binding, DataSource, DocentricTemplate, SchemaNode, TemplateElement } from '../model/types.js';

/**
 * Resolves every binding of every element against the data source schemas.
 *
 * Context rules (as the Docentric engine applies them):
 *   - Source=<data source name>: paths start at that data source's root element
 *     (ReportData, GeneralData, Parameters), so `@IsWithAmount` is an attribute of the root.
 *   - Source='.': paths start at the current item of the nearest enclosing List/Group.
 *   - SummaryField AggregateBinding: paths start at the collection its ValueBinding selects.
 */

const CONTEXT_PROVIDERS = new Set(['List', 'Group']);

export interface SchemaContext {
  dataSource: DataSource;
  /** Root first, current node last. */
  chain: SchemaNode[];
}

export type Resolution =
  | { status: 'ok'; chain: SchemaNode[]; attribute: string | null }
  | { status: 'missing'; resolvedText: string; step: string; kind: 'element' | 'attribute'; suggestions: string[] }
  | { status: 'skipped'; reason: string };

export interface PathCheck {
  ref: PathRef;
  resolution: Resolution;
}

export interface BoundExpression {
  binding: Binding;
  info: ExpressionInfo;
  context: SchemaContext | null;
  /** Why no context could be established (unknown source, '.' outside a List, ...). */
  contextError: { code: 'unknown-source' | 'no-context' | 'no-schema'; message: string } | null;
  paths: PathCheck[];
  /** When the whole expression is a single element path: the node it selects (a List's item type). */
  selects: SchemaContext | null;
}

export interface ElementAnalysis {
  element: TemplateElement;
  parentId: string | null;
  value: BoundExpression | null;
  aggregate: BoundExpression | null;
}

export function analyzeBindings(template: DocentricTemplate): Map<string, ElementAnalysis> {
  const parents = buildParentMap(template);
  const byName = new Map(template.dataSources.map((d) => [d.name, d]));
  const out = new Map<string, ElementAnalysis>();
  const listContext = new Map<string, SchemaContext | null>();

  /** The item context an element sees through its nearest List/Group ancestor. */
  const enclosingContext = (id: string, seen: Set<string>): SchemaContext | null => {
    let cur = parents.get(id) ?? null;
    while (cur) {
      if (seen.has(cur)) return null;
      seen.add(cur);
      const el = template.elements.get(cur);
      if (el && CONTEXT_PROVIDERS.has(el.type)) {
        if (!listContext.has(cur)) listContext.set(cur, analyze(el, seen).value?.selects ?? null);
        return listContext.get(cur) ?? null;
      }
      cur = parents.get(cur) ?? null;
    }
    return null;
  };

  const bind = (
    binding: Binding,
    el: TemplateElement,
    seen: Set<string>,
    base?: SchemaContext | null,
  ): BoundExpression => {
    const info = analyzeExpression(binding.path);
    const result: BoundExpression = { binding, info, context: null, contextError: null, paths: [], selects: null };

    if (base !== undefined) {
      result.context = base;
      if (!base)
        result.contextError = { code: 'no-context', message: 'the collection it aggregates could not be resolved' };
    } else if (!binding.source || binding.source === '.') {
      result.context = enclosingContext(el.id, seen);
      if (!result.context) {
        result.contextError = {
          code: 'no-context',
          message: "Source='.' but no enclosing List/Group with a resolvable collection",
        };
      }
    } else {
      const ds = byName.get(binding.source);
      if (!ds) {
        result.contextError = { code: 'unknown-source', message: `unknown data source '${binding.source}'` };
      } else if (!ds.root) {
        result.contextError = { code: 'no-schema', message: `data source '${ds.name}' has no XSD` };
      } else {
        result.context = { dataSource: ds, chain: [ds.root] };
      }
    }

    if (result.context) {
      const ctx = result.context;
      result.paths = info.paths.map((ref) => ({ ref, resolution: resolvePath(ref, ctx) }));
      const only = result.paths.length === 1 ? result.paths[0] : undefined;
      if (
        only &&
        only.resolution.status === 'ok' &&
        !only.resolution.attribute &&
        only.ref.text === binding.path.trim()
      ) {
        result.selects = { dataSource: ctx.dataSource, chain: only.resolution.chain };
      }
    }
    return result;
  };

  const analyze = (el: TemplateElement, seen: Set<string> = new Set([el.id])): ElementAnalysis => {
    const cached = out.get(el.id);
    if (cached) return cached;
    const value = el.binding ? bind(el.binding, el, seen) : null;
    const aggregate = el.aggregateBinding ? bind(el.aggregateBinding, el, seen, value?.selects ?? null) : null;
    const record: ElementAnalysis = { element: el, parentId: parents.get(el.id) ?? null, value, aggregate };
    out.set(el.id, record);
    return record;
  };

  for (const el of template.elements.values()) analyze(el);
  return out;
}

/**
 * Parent of each element: the ChildElements registry first (it is what the engine
 * uses), the content-control nesting in the document as a fallback.
 */
export function buildParentMap(template: DocentricTemplate): Map<string, string> {
  const parents = new Map<string, string>();
  for (const el of template.elements.values()) {
    for (const childId of el.childIds) if (!parents.has(childId)) parents.set(childId, el.id);
  }
  for (const cc of template.contentControls) {
    if (cc.elementId && cc.parentElementId && !parents.has(cc.elementId)) parents.set(cc.elementId, cc.parentElementId);
  }
  return parents;
}

export function resolvePath(ref: PathRef, ctx: SchemaContext): Resolution {
  if (ref.descendant) return { status: 'skipped', reason: "uses '//' (descendant axis)" };
  const root = ctx.chain[0] as SchemaNode;
  let chain: SchemaNode[] = ref.absolute ? [] : [...ctx.chain];
  const done: string[] = [];
  let attribute: string | null = null;

  for (const step of ref.steps) {
    if (attribute) {
      return {
        status: 'missing',
        resolvedText: done.join('/'),
        step: stepText(step),
        kind: 'element',
        suggestions: [],
      };
    }
    const cur = chain[chain.length - 1];
    switch (step.kind) {
      case 'wildcard':
        return { status: 'skipped', reason: 'uses a wildcard step' };
      case 'self':
        break;
      case 'parent':
        if (chain.length <= 1) {
          return { status: 'missing', resolvedText: done.join('/'), step: '..', kind: 'element', suggestions: [] };
        }
        chain = chain.slice(0, -1);
        break;
      case 'element': {
        const next = cur ? cur.children.get(step.name) : step.name === root.name ? root : undefined;
        if (!next) {
          const pool = cur ? [...cur.children.keys()] : [root.name];
          return {
            status: 'missing',
            resolvedText: done.join('/'),
            step: step.name,
            kind: 'element',
            suggestions: closest(step.name, pool),
          };
        }
        chain = [...chain, next];
        break;
      }
      case 'attribute': {
        if (!cur?.attributes.has(step.name)) {
          return {
            status: 'missing',
            resolvedText: done.join('/'),
            step: `@${step.name}`,
            kind: 'attribute',
            suggestions: cur ? closest(step.name, [...cur.attributes.keys()]).map((s) => `@${s}`) : [],
          };
        }
        attribute = step.name;
        break;
      }
    }
    done.push(stepText(step));
  }
  if (chain.length === 0) return { status: 'skipped', reason: 'bare root path' };
  return { status: 'ok', chain, attribute };
}

function stepText(step: PathRef['steps'][number]): string {
  switch (step.kind) {
    case 'element':
      return step.name;
    case 'attribute':
      return `@${step.name}`;
    case 'self':
      return '.';
    case 'parent':
      return '..';
    case 'wildcard':
      return '*';
  }
}

/** Up to three candidates, closest first, by case-insensitive edit distance and prefix/substring match. */
export function closest(target: string, pool: string[]): string[] {
  const t = target.toLowerCase();
  return pool
    .map((name) => {
      const n = name.toLowerCase();
      let score = levenshtein(t, n);
      // A substring hit (Price → SalesPrice, Company → CurrentCompany) always qualifies.
      if (n.includes(t) || t.includes(n)) score = Math.min(score, 1 + Math.abs(n.length - t.length) / 100);
      return { name, score };
    })
    .filter((c) => c.score <= Math.max(2, Math.floor(target.length / 3)))
    .sort((a, b) => a.score - b.score || a.name.localeCompare(b.name))
    .slice(0, 3)
    .map((c) => c.name);
}

function levenshtein(a: string, b: string): number {
  const row = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    let prev = row[0] as number;
    row[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = row[j] as number;
      row[j] = Math.min((row[j] as number) + 1, (row[j - 1] as number) + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = tmp;
    }
  }
  return row[b.length] as number;
}
