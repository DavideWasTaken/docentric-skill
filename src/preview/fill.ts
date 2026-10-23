import type { Document, Element, Node } from '@xmldom/xmldom';
import { contentParts, findProjectDataPart, loadTemplate } from '../docx/loadTemplate.js';
import { NS, WordPackage } from '../docx/package.js';
import {
  asNodes,
  type EvalContext,
  evaluateSource,
  formatPlain,
  formatValue,
  toBoolean,
  toNumber,
  toStringValue,
  type Value,
  withLocale,
  type XNode,
} from '../expr/evaluate.js';
import type { DocentricTemplate, TemplateElement } from '../model/types.js';
import { child, children, localName, path } from '../xml/dom.js';

/**
 * Produces a filled copy of a template: every Docentric content control is
 * replaced by what it would print with the template's own sample data —
 * Fields and Labels get their text, Lists repeat their rows, Ifs keep or drop
 * their content, Vars are computed in document order. The result is a plain
 * Word document (the ProjectData part is removed) ready to be converted to PDF.
 *
 * It approximates the Docentric engine; `issues` lists what it could not do.
 */

const DOCENTRIC_TAG = /^\d+-(\d+)$/;
/** Confirmed on real templates: AggregateFunction="1" is Sum. Other codes are not documented. */
const AGGREGATE_SUM = '1';

export interface FillOptions {
  /** Template language to use for labels and number/date formatting. Default: the template default. */
  language?: string;
}

export interface FillResult {
  pkg: WordPackage;
  language: string;
  locale: string;
  stats: { fields: number; labels: number; lists: number; listItems: number; ifsTrue: number; ifsFalse: number };
  issues: string[];
}

export async function fillTemplate(data: Buffer | Uint8Array, opts: FillOptions = {}): Promise<FillResult> {
  const pkg = await WordPackage.open(data);
  const template = await loadTemplate(pkg);
  const pdPart = (await findProjectDataPart(pkg)) as string;
  const pd = (await pkg.dom(pdPart)) as Document;

  const language = opts.language ?? template.defaultLanguage ?? template.languages[0] ?? 'en-us';
  const locale = toLocale(language);
  const roots = sampleRoots(pd);
  const issues = new Set<string>();
  const unsupported = new Set<string>();
  const vars = withLocale(new Map<string, Value>(), locale);
  const stats = { fields: 0, labels: 0, lists: 0, listItems: 0, ifsTrue: 0, ifsFalse: 0 };

  if (roots.size === 0) issues.add('The template has no sample data: fields will be empty.');

  const labelText = labelResolver(template, language);
  const filler = new Filler(template, roots, vars, labelText, locale, stats, issues, unsupported);

  for (const part of contentParts(pkg)) {
    const doc = await pkg.dom(part);
    if (!doc) continue;
    filler.processChildren(doc.documentElement as Element, null);
    ensureCellParagraphs(doc);
    pkg.setDom(part, doc);
  }

  await removeProjectData(pkg, pdPart);
  for (const u of unsupported) issues.add(`Not evaluated (shown as a marker): ${u}`);
  return { pkg, language, locale, stats, issues: [...issues] };
}

class Filler {
  constructor(
    private readonly template: DocentricTemplate,
    private readonly roots: Map<string, Element>,
    private readonly vars: Map<string, Value>,
    private readonly labelText: (code: string) => string | null,
    private readonly locale: string,
    private readonly stats: FillResult['stats'],
    private readonly issues: Set<string>,
    private readonly unsupported: Set<string>,
  ) {}

  /** Walks a container; `item` is the current List item (for Source='.'). */
  processChildren(container: Element, item: XNode | null): void {
    for (const node of snapshot(container)) this.processNode(node, item);
  }

  processNode(node: Element, item: XNode | null): void {
    if (localName(node) !== 'sdt') {
      this.processChildren(node, item);
      return;
    }
    const el = this.elementOf(node);
    const content = child(node, 'sdtContent');
    if (!el || !content) {
      if (content) this.processChildren(content, item);
      return;
    }
    switch (el.type) {
      case 'If': {
        const keep = this.condition(el, item);
        if (keep) {
          this.stats.ifsTrue++;
          this.processChildren(content, item);
          unwrap(node);
        } else {
          this.stats.ifsFalse++;
          removeConditional(node);
        }
        return;
      }
      case 'List':
      case 'Group': {
        if (el.type === 'Group') this.issues.add('Group elements are previewed like a List (no grouping).');
        const items = el.binding ? asNodes(this.eval(el, el.binding.source, el.binding.path, item)) : [];
        this.stats.lists++;
        this.stats.listItems += items.length;
        const parent = node.parentNode as Element;
        for (const it of items) {
          for (const c of snapshot(content)) {
            const clone = c.cloneNode(true) as Element;
            parent.insertBefore(clone, node);
            this.processNode(clone, it);
          }
        }
        parent.removeChild(node);
        return;
      }
      case 'Field': {
        this.stats.fields++;
        const value = el.binding ? this.eval(el, el.binding.source, el.binding.path, item) : '';
        setText(content, this.display(value, el.attributes.FormatString));
        unwrap(node);
        return;
      }
      case 'Label': {
        this.stats.labels++;
        const code = el.labelCode ?? '';
        setText(content, this.labelText(code) ?? code);
        unwrap(node);
        return;
      }
      case 'SummaryField': {
        this.stats.fields++;
        const rows = el.binding ? asNodes(this.eval(el, el.binding.source, el.binding.path, item)) : [];
        const fn = el.attributes.AggregateFunction ?? AGGREGATE_SUM;
        if (fn !== AGGREGATE_SUM) this.issues.add(`SummaryField ${el.id}: AggregateFunction=${fn} previewed as Sum.`);
        let total = 0;
        for (const row of rows) {
          total += el.aggregateBinding ? toNumber(this.eval(el, '.', el.aggregateBinding.path, row)) : 1;
        }
        setText(content, this.display(total, el.attributes.FormatString));
        unwrap(node);
        return;
      }
      case 'Var': {
        if (el.name && el.binding) this.vars.set(el.name, this.eval(el, el.binding.source, el.binding.path, item));
        removeConditional(node);
        return;
      }
      case 'Image':
        this.issues.add('Images keep their design-time placeholder.');
        unwrap(node);
        return;
      default:
        this.issues.add(`Element type ${el.type} is not previewed; its design-time content is shown.`);
        this.processChildren(content, item);
        unwrap(node);
    }
  }

  private elementOf(sdt: Element): TemplateElement | null {
    const pr = child(sdt, 'sdtPr');
    const tag = pr ? child(pr, 'tag') : null;
    const m = tag ? DOCENTRIC_TAG.exec(tag.getAttribute('w:val') ?? '') : null;
    return m ? (this.template.elements.get(m[1] as string) ?? null) : null;
  }

  private condition(el: TemplateElement, item: XNode | null): boolean {
    if (!el.binding) return true;
    const v = toBoolean(this.eval(el, el.binding.source, el.binding.path, item));
    const invert = el.attributes.Invert === '1' || el.attributes.Invert?.toLowerCase() === 'true';
    return invert ? !v : v;
  }

  private eval(el: TemplateElement, source: string | null, expr: string, item: XNode | null): Value {
    const fromItem = !source || source === '.';
    const root = fromItem ? rootOf(item) : (this.roots.get(source) ?? null);
    const ctx: EvalContext = {
      node: fromItem ? item : root,
      root,
      vars: this.vars,
      label: this.labelText,
      unsupported: this.unsupported,
    };
    if (!fromItem && !root) {
      this.issues.add(`No sample data for data source '${source}'.`);
      return '';
    }
    try {
      return evaluateSource(expr, ctx);
    } catch (e) {
      this.issues.add(`${el.type} ${el.id}: ${(e as Error).message}`);
      return `⟨error⟩`;
    }
  }

  private display(v: Value, format: string | undefined): string {
    if (typeof v === 'number') return format ? formatValue(v, format, this.locale) : formatPlain(v, this.locale);
    const s = toStringValue(v);
    if (format) return formatValue(v, format, this.locale);
    // Sample amounts come as xs:decimal text ("3517.7800000000"): show them as the culture would.
    if (/^-?\d+\.\d+$/.test(s)) return formatPlain(Number(s), this.locale);
    if (/^\d{4}-\d{2}-\d{2}(T00:00:00(\.0+)?)?$/.test(s)) return formatValue(s, 'd', this.locale);
    return s;
  }
}

// --- Sample data and labels -------------------------------------------------------

function sampleRoots(pd: Document): Map<string, Element> {
  const out = new Map<string, Element>();
  const container = child(pd.documentElement as Element, 'DataSources');
  if (!container) return out;
  for (const ds of children(container, 'DataSource')) {
    const sample = path(ds, 'PreviewData', 'SampleXml');
    const root = sample ? children(sample)[0] : undefined;
    const name = ds.getAttribute('Name');
    if (root && name) out.set(name, root);
  }
  return out;
}

function rootOf(item: XNode | null): Element | null {
  if (!item) return null;
  let cur: Node | null =
    item.nodeType === 2 ? ((item as { ownerElement?: Element | null }).ownerElement ?? null) : item;
  while (cur?.parentNode && cur.parentNode.nodeType === 1 && localName(cur.parentNode as Element) !== 'SampleXml') {
    cur = cur.parentNode;
  }
  return (cur as Element | null) ?? null;
}

function labelResolver(template: DocentricTemplate, language: string): (code: string) => string | null {
  const byCode = new Map(template.labels.map((l) => [l.code, l.texts]));
  const fallback = template.defaultLanguage;
  return (code) => {
    const texts = byCode.get(code);
    if (!texts) return null;
    return texts[language] ?? (fallback ? texts[fallback] : undefined) ?? Object.values(texts)[0] ?? null;
  };
}

export function toLocale(language: string): string {
  const l = language.toLowerCase();
  if (l.includes('-')) {
    const [a, b] = l.split('-');
    return `${a}-${(b ?? '').toUpperCase()}`;
  }
  const defaults: Record<string, string> = { it: 'it-IT', en: 'en-US', de: 'de-DE', fr: 'fr-FR', es: 'es-ES' };
  return defaults[l] ?? l;
}

// --- WordprocessingML surgery -----------------------------------------------------

function snapshot(container: Element): Element[] {
  return children(container);
}

/** Replaces an sdt with its content. */
function unwrap(sdt: Element): void {
  const parent = sdt.parentNode as Element;
  const content = child(sdt, 'sdtContent');
  if (content) for (const c of snapshot(content)) parent.insertBefore(c, sdt);
  parent.removeChild(sdt);
}

/**
 * Removes an sdt whose condition is false. Cell-level content cannot vanish
 * (the table grid would break), so cells are kept and emptied; rows and blocks
 * are removed.
 */
function removeConditional(sdt: Element): void {
  const parent = sdt.parentNode as Element;
  const content = child(sdt, 'sdtContent');
  for (const c of content ? snapshot(content) : []) {
    if (localName(c) === 'tc') {
      for (const inner of snapshot(c)) if (localName(inner) !== 'tcPr') c.removeChild(inner);
      c.appendChild((c.ownerDocument as Document).createElementNS(NS.w, 'w:p'));
      parent.insertBefore(c, sdt);
    }
  }
  parent.removeChild(sdt);
}

/** Writes `text` into the first w:t of the content and clears the others. Newlines become w:br. */
function setText(content: Element, text: string): void {
  const doc = content.ownerDocument as Document;
  const ts = Array.from(content.getElementsByTagNameNS(NS.w, 't'));
  let first = ts[0];
  if (!first) {
    const p = Array.from(content.getElementsByTagNameNS(NS.w, 'p'))[0];
    const runHost = p ?? (localName(content) === 'sdtContent' ? content : null);
    if (!runHost) return;
    const r = doc.createElementNS(NS.w, 'w:r');
    first = doc.createElementNS(NS.w, 'w:t');
    r.appendChild(first);
    runHost.appendChild(r);
  }
  for (const t of ts.slice(1)) t.parentNode?.removeChild(t);
  for (const pe of Array.from(content.getElementsByTagNameNS(NS.w, 'proofErr'))) pe.parentNode?.removeChild(pe);

  const lines = text.split('\n');
  while (first.firstChild) first.removeChild(first.firstChild);
  first.setAttribute('xml:space', 'preserve');
  first.appendChild(doc.createTextNode(lines[0] ?? ''));
  let after: Element = first;
  for (const line of lines.slice(1)) {
    const br = doc.createElementNS(NS.w, 'w:br');
    const t = doc.createElementNS(NS.w, 'w:t');
    t.setAttribute('xml:space', 'preserve');
    t.appendChild(doc.createTextNode(line));
    insertAfter(br, after);
    insertAfter(t, br);
    after = t;
  }
}

function insertAfter(node: Element, ref: Element): void {
  const parent = ref.parentNode as Element;
  if (ref.nextSibling) parent.insertBefore(node, ref.nextSibling);
  else parent.appendChild(node);
}

/** Word refuses a table cell without a paragraph; removals can leave one empty. */
function ensureCellParagraphs(doc: Document): void {
  for (const tc of Array.from(doc.getElementsByTagNameNS(NS.w, 'tc'))) {
    if (!children(tc).some((c) => localName(c) === 'p' || localName(c) === 'tbl')) {
      tc.appendChild(doc.createElementNS(NS.w, 'w:p'));
    }
  }
}

/** The preview is a plain document: drop the Docentric part so nobody mistakes it for a template. */
async function removeProjectData(pkg: WordPackage, part: string): Promise<void> {
  const item = part.split('/').pop() as string;
  pkg.remove(part);
  pkg.remove(`customXml/_rels/${item}.rels`);
  const relsDoc = await pkg.dom('word/_rels/document.xml.rels');
  if (relsDoc) {
    for (const r of Array.from((relsDoc.documentElement as Element).getElementsByTagName('Relationship'))) {
      if ((r.getAttribute('Target') ?? '').endsWith(`customXml/${item}`)) r.parentNode?.removeChild(r);
    }
    pkg.setDom('word/_rels/document.xml.rels', relsDoc);
  }
}
