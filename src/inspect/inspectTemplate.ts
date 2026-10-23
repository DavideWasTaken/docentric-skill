import { analyzeBindings, type ElementAnalysis } from '../analysis/bindings.js';
import { isTemplateLabel } from '../model/projectData.js';
import type { DocentricTemplate, SchemaNode } from '../model/types.js';
import { countSchema } from '../schema/xsd.js';

export type InspectSection = 'summary' | 'elements' | 'schema' | 'labels';

export interface InspectOptions {
  include?: InspectSection[];
  /** Restrict the schema section to one data source. */
  dataSource?: string;
  /** Levels of the schema tree to print below each root (default 2). */
  schemaDepth?: number;
}

/**
 * Text rendering meant for an LLM: compact, one line per element, with the data
 * context each binding is evaluated in, so the agent can write new paths without
 * opening Word.
 */
export function inspectTemplate(template: DocentricTemplate, opts: InspectOptions = {}): string {
  const include = new Set<InspectSection>(opts.include?.length ? opts.include : ['summary', 'elements']);
  const out: string[] = [];
  if (include.has('summary')) out.push(renderSummary(template));
  if (include.has('elements')) out.push(renderElements(template));
  if (include.has('schema')) out.push(renderSchema(template, opts.dataSource, opts.schemaDepth ?? 2));
  if (include.has('labels')) out.push(renderLabels(template));
  return out.join('\n\n');
}

function renderSummary(t: DocentricTemplate): string {
  const typeCounts = new Map<string, number>();
  for (const e of t.elements.values()) typeCounts.set(e.type, (typeCounts.get(e.type) ?? 0) + 1);
  const parts = new Map<string, number>();
  for (const c of t.contentControls) parts.set(c.part, (parts.get(c.part) ?? 0) + 1);
  const lines = [
    '## Template',
    `Docentric version: ${t.productVersion ?? 'unknown'}`,
    `Report: ${t.report.reportId ?? '—'} (data source type ${t.report.dataSourceType ?? '—'}, group ${t.report.reportGroupId ?? '—'})`,
    `Data source provider class: ${t.report.dataSourceProviderClass ?? '—'}`,
    `Languages: ${t.languages.join(', ') || '—'} (default ${t.defaultLanguage ?? '—'})`,
    `Elements: ${t.elements.size} — ${[...typeCounts].map(([k, v]) => `${v} ${k}`).join(', ')}`,
    `Placed in: ${[...parts].map(([p, n]) => `${p} (${n})`).join(', ') || 'no content controls'}`,
    `Template labels: ${t.labels.filter(isTemplateLabel).length} (+ ${t.labels.filter((l) => !isTemplateLabel(l)).length} cached D365/ISV label texts)`,
    'Data sources:',
  ];
  for (const ds of t.dataSources) {
    const c = ds.root ? countSchema(ds.root) : null;
    lines.push(
      `- ${ds.name} [${ds.section}] root=${ds.root?.name ?? '—'}` +
        (c ? `, ${c.elements} elements, ${c.attributes} attributes` : ', no XSD') +
        (ds.hasSampleXml ? ', sample data embedded' : ''),
    );
  }
  return lines.join('\n');
}

function renderElements(t: DocentricTemplate): string {
  const analysis = analyzeBindings(t);
  const childrenOf = new Map<string | null, string[]>();
  // Document order first, so the tree reads like the page.
  const order = new Map<string, number>();
  t.contentControls.forEach((c, i) => {
    if (c.elementId && !order.has(c.elementId)) order.set(c.elementId, i);
  });
  const ids = [...t.elements.keys()].sort((a, b) => (order.get(a) ?? 1e9) - (order.get(b) ?? 1e9));
  for (const id of ids) {
    const parent = analysis.get(id)?.parentId ?? null;
    const list = childrenOf.get(parent) ?? [];
    list.push(id);
    childrenOf.set(parent, list);
  }
  const partOf = new Map(t.contentControls.map((c) => [c.elementId, c.part]));

  const lines = ['## Elements (tree from ChildElements, document order)'];
  const visited = new Set<string>();
  const walk = (id: string, depth: number, part: string | undefined) => {
    if (visited.has(id)) return;
    visited.add(id);
    const a = analysis.get(id);
    if (!a) return;
    const myPart = partOf.get(id);
    const partNote = myPart && myPart !== part ? ` {${myPart.replace('word/', '')}}` : '';
    lines.push(`${'  '.repeat(depth)}- ${describe(a)}${partNote}`);
    for (const c of childrenOf.get(id) ?? []) walk(c, depth + 1, myPart ?? part);
  };
  for (const id of childrenOf.get(null) ?? []) walk(id, 0, undefined);
  return lines.join('\n');
}

function describe(a: ElementAnalysis): string {
  const e = a.element;
  let s = `${e.type}#${e.id}`;
  if (e.name) s += ` '${e.name}'`;
  if (e.labelCode) s += ` ${e.labelCode}`;
  if (e.binding) {
    const ctx = a.value?.context;
    const src =
      e.binding.source === '.' ? `.(${ctx ? ctx.chain.map((n) => n.name).join('/') : '?'})` : e.binding.source;
    s += ` ← ${src}: ${e.binding.path}`;
  }
  if (e.aggregateBinding) s += ` aggregate(fn=${e.attributes.AggregateFunction ?? '?'}): ${e.aggregateBinding.path}`;
  const extras = ['FormatString', 'RenderMode', 'Invert']
    .filter((k) => e.attributes[k])
    .map((k) => `${k}=${e.attributes[k]}`);
  if (extras.length) s += ` [${extras.join(', ')}]`;
  return s;
}

function renderSchema(t: DocentricTemplate, only: string | undefined, depth: number): string {
  const lines = ['## Data source schemas (paths are relative to the root element)'];
  const sources = t.dataSources.filter((d) => !only || d.name.toLowerCase() === only.toLowerCase());
  if (only && sources.length === 0) {
    lines.push(`No data source named '${only}'. Available: ${t.dataSources.map((d) => d.name).join(', ')}`);
  }
  for (const ds of sources) {
    lines.push(`### ${ds.name} [${ds.section}]`);
    if (!ds.root) {
      lines.push('(no XSD)');
      continue;
    }
    renderNode(ds.root, 0, depth, lines);
  }
  return lines.join('\n');
}

function renderNode(node: SchemaNode, level: number, maxDepth: number, lines: string[]): void {
  const pad = '  '.repeat(level);
  const attrs = [...node.attributes].map(([n, ty]) => `@${n}:${ty}`).join(' ');
  lines.push(`${pad}${node.name}${node.repeated ? '[*]' : ''}${attrs ? `  ${attrs}` : ''}`);
  if (level >= maxDepth) {
    if (node.children.size) lines.push(`${pad}  … ${node.children.size} child element(s), raise schemaDepth to expand`);
    return;
  }
  for (const c of node.children.values()) renderNode(c, level + 1, maxDepth, lines);
}

function renderLabels(t: DocentricTemplate): string {
  const lines = ['## Template labels'];
  const cached: string[] = [];
  for (const l of t.labels) {
    if (!isTemplateLabel(l)) {
      cached.push(l.code);
      continue;
    }
    const texts = Object.entries(l.texts)
      .map(([lang, v]) => `${lang}: ${JSON.stringify(v.length > 120 ? `${v.slice(0, 117)}...` : v)}`)
      .join(' | ');
    lines.push(`- ${l.code} [${l.section}] ${texts}`);
  }
  if (cached.length) {
    lines.push(
      `Cached D365/ISV label texts (${cached.length}, resolved again at print time — look them up with d365fo-mcp): ` +
        `${cached.slice(0, 20).join(', ')}${cached.length > 20 ? ', …' : ''}`,
    );
  }
  return lines.join('\n');
}
