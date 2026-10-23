import { type Document, DOMParser, type Element, type Node } from '@xmldom/xmldom';
import { NS } from '../docx/package.js';
import { elementXml, type LabelDef, labelXml, type NewElement, TEMPLATES_NS } from '../write/projectDataXml.js';
import { DOC_NAMESPACES } from '../write/assemble.js';
import { attr, child, children, localName } from '../xml/dom.js';

/**
 * In-place edits of an existing template: the ProjectData registry and the
 * WordprocessingML parts are changed node by node, so everything this tool does
 * not understand (styles, fields, tracked changes, Designer settings) survives.
 */

// --- ProjectData ---------------------------------------------------------------

export class Registry {
  constructor(readonly doc: Document) {}

  private get root(): Element {
    return this.doc.documentElement as Element;
  }

  private get elementsContainer(): Element {
    let c = child(this.root, 'Elements');
    if (!c) {
      c = this.doc.createElementNS(null, 'Elements');
      this.root.appendChild(c);
    }
    return c;
  }

  element(id: string): Element | null {
    return children(this.elementsContainer, 'Element').find((e) => attr(e, 'Id') === id) ?? null;
  }

  ids(): string[] {
    return children(this.elementsContainer, 'Element').map((e) => attr(e, 'Id') ?? '');
  }

  add(e: NewElement): void {
    for (const node of parseFragment(elementXml(e), ''))
      this.elementsContainer.appendChild(this.doc.importNode(node, true));
  }

  childIds(id: string): string[] {
    const el = this.element(id);
    const kids = el ? child(el, 'ChildElements') : null;
    return kids ? children(kids, 'ElementRef').map((r) => attr(r, 'Id') ?? '') : [];
  }

  parentOf(id: string): string | null {
    for (const e of children(this.elementsContainer, 'Element')) {
      const kids = child(e, 'ChildElements');
      if (kids && children(kids, 'ElementRef').some((r) => attr(r, 'Id') === id)) return attr(e, 'Id');
    }
    return null;
  }

  /** Inserts `childId` into `parentId`'s ChildElements, after `afterId` when given (else at the end). */
  addChild(parentId: string, childId: string, afterId?: string | null): void {
    const el = this.element(parentId);
    if (!el) return;
    let kids = child(el, 'ChildElements');
    if (!kids) {
      kids = this.doc.createElementNS(null, 'ChildElements');
      el.insertBefore(kids, el.firstChild);
    }
    const ref = this.doc.createElementNS(null, 'ElementRef');
    ref.setAttribute('Id', childId);
    const after = afterId ? children(kids, 'ElementRef').find((r) => attr(r, 'Id') === afterId) : undefined;
    if (after?.nextSibling) kids.insertBefore(ref, after.nextSibling);
    else kids.appendChild(ref);
  }

  removeChild(parentId: string, childId: string): void {
    const el = this.element(parentId);
    const kids = el ? child(el, 'ChildElements') : null;
    if (!kids) return;
    for (const r of children(kids, 'ElementRef')) if (attr(r, 'Id') === childId) kids.removeChild(r);
    if (children(kids).length === 0) el?.removeChild(kids);
  }

  /** Removes an element and, recursively, the elements it contains. Returns the removed ids. */
  remove(id: string): string[] {
    const removed: string[] = [];
    const walk = (x: string) => {
      for (const c of this.childIds(x)) walk(c);
      const el = this.element(x);
      if (el) {
        this.elementsContainer.removeChild(el);
        removed.push(x);
      }
    };
    const parent = this.parentOf(id);
    if (parent) this.removeChild(parent, id);
    walk(id);
    return removed;
  }

  binding(id: string): Element | null {
    const el = this.element(id);
    return el ? child(el, 'ValueBinding') : null;
  }

  setAttribute(id: string, name: string, value: string | null): void {
    const el = this.element(id);
    if (!el) return;
    if (value === null) el.removeAttribute(name);
    else el.setAttribute(name, value);
  }

  // Labels live in the dataSourcePackage, in the templates namespace with a file-specific prefix.
  private get labelsContainer(): Element | null {
    const pkg = child(this.root, 'dataSourcePackage');
    return pkg ? child(pkg, 'labels') : null;
  }

  label(code: string): Element | null {
    const labels = this.labelsContainer;
    return labels ? (children(labels, 'label').find((l) => attr(l, 'code') === code) ?? null) : null;
  }

  upsertLabel(def: LabelDef): 'added' | 'updated' {
    const labels = this.labelsContainer;
    if (!labels) throw new Error('This template has no labels block (dataSourcePackage/labels).');
    const prefix = templatesPrefix(labels) ?? 'd2p1';
    const existing = this.label(def.code);
    if (existing) {
      for (const [lang, text] of Object.entries(def.texts)) {
        const t = children(existing, 'text').find((x) => attr(x, 'language') === lang);
        if (t) t.setAttributeNS(TEMPLATES_NS, `${prefix}:value`, text);
        else {
          const n = this.doc.createElementNS(TEMPLATES_NS, 'text');
          n.setAttributeNS(TEMPLATES_NS, `${prefix}:language`, lang);
          n.setAttributeNS(TEMPLATES_NS, `${prefix}:value`, text);
          existing.appendChild(n);
        }
      }
      return 'updated';
    }
    for (const node of parseFragment(
      labelXml(def, prefix),
      `xmlns="${TEMPLATES_NS}" xmlns:${prefix}="${TEMPLATES_NS}"`,
    )) {
      labels.appendChild(this.doc.importNode(node, true));
    }
    return 'added';
  }

  /** Adds a template language; returns false when it is already there. */
  addLanguage(language: string, makeDefault: boolean): boolean {
    const pkg = child(this.root, 'dataSourcePackage');
    const langs = pkg ? child(pkg, 'languages') : null;
    if (!langs) throw new Error('This template has no languages block (dataSourcePackage/languages).');
    const prefix = templatesPrefix(langs) ?? 'd2p1';
    const exists = this.languages().includes(language);
    if (!exists) {
      const el = this.doc.createElementNS(TEMPLATES_NS, 'language');
      el.setAttributeNS(TEMPLATES_NS, `${prefix}:value`, language);
      langs.appendChild(el);
    }
    if (makeDefault) langs.setAttributeNS(TEMPLATES_NS, `${prefix}:defaultLanguage`, language);
    return !exists;
  }

  languages(): string[] {
    const pkg = child(this.root, 'dataSourcePackage');
    const langs = pkg ? child(pkg, 'languages') : null;
    return langs
      ? children(langs, 'language')
          .map((l) => attr(l, 'value') ?? '')
          .filter(Boolean)
      : [];
  }
}

/** The prefix this file uses for the templates namespace (d2p1 in Designer files, others elsewhere). */
function templatesPrefix(scope: Element): string | null {
  for (let n: Node | null = scope; n && n.nodeType === 1; n = n.parentNode) {
    const el = n as Element;
    for (const l of children(el)) {
      const attrs = l.attributes;
      for (let i = 0; i < attrs.length; i++) {
        const a = attrs.item(i);
        if (a?.namespaceURI === TEMPLATES_NS && a.prefix) return a.prefix;
      }
    }
    const attrs = el.attributes;
    for (let i = 0; i < attrs.length; i++) {
      const a = attrs.item(i);
      if (a?.name.startsWith('xmlns:') && a.value === TEMPLATES_NS) return a.name.slice(6);
    }
  }
  return null;
}

// --- WordprocessingML -----------------------------------------------------------

/** Parses an XML fragment (several sibling elements) with the given namespace declarations. */
export function parseFragment(xml: string, namespaces: string): Element[] {
  const doc = new DOMParser().parseFromString(`<fragment ${namespaces}>${xml}</fragment>`, 'text/xml');
  return children(doc.documentElement as Element);
}

export function parseWordFragment(xml: string): Element[] {
  return parseFragment(xml, DOC_NAMESPACES.replace(/mc:Ignorable="[^"]*"/, ''));
}

export function insertNodes(target: Document, nodes: Element[], parent: Node, before: Node | null): Element[] {
  const out: Element[] = [];
  for (const n of nodes) {
    const imported = target.importNode(n, true) as Element;
    if (before) parent.insertBefore(imported, before);
    else parent.appendChild(imported);
    out.push(imported);
  }
  return out;
}

const TAG = /^(\d+)-(\d+)$/;

export function sdtTag(sdt: Element): { projectId: string; elementId: string } | null {
  const pr = child(sdt, 'sdtPr');
  const tag = pr ? child(pr, 'tag') : null;
  const m = tag ? TAG.exec(tag.getAttributeNS(NS.w, 'val') ?? tag.getAttribute('w:val') ?? '') : null;
  return m ? { projectId: m[1] as string, elementId: m[2] as string } : null;
}

export function findSdt(doc: Document, elementId: string): Element | null {
  for (const s of Array.from(doc.getElementsByTagNameNS(NS.w, 'sdt'))) {
    if (sdtTag(s)?.elementId === elementId) return s;
  }
  return null;
}

/** Nearest enclosing Docentric element of a node (its sdt ancestor), or null at top level. */
export function enclosingElementId(node: Node): string | null {
  for (let n: Node | null = node.parentNode; n; n = n.parentNode) {
    if (n.nodeType === 1 && localName(n as Element) === 'sdt') {
      const t = sdtTag(n as Element);
      if (t) return t.elementId;
    }
  }
  return null;
}

export function isInside(node: Node, container: Node): boolean {
  for (let n: Node | null = node.parentNode; n; n = n.parentNode) if (n === container) return true;
  return false;
}

export function ancestor(node: Node, name: string): Element | null {
  for (let n: Node | null = node.parentNode; n; n = n.parentNode) {
    if (n.nodeType === 1 && localName(n as Element) === name) return n as Element;
  }
  return null;
}

/** First element child of an sdt's content: tells whether it wraps runs, paragraphs, cells or rows. */
export function sdtLevel(sdt: Element): 'run' | 'block' | 'cell' | 'row' {
  // Look through nested content controls (Designer wraps rows in untagged repeating sections).
  let first: Element | undefined = children(child(sdt, 'sdtContent') ?? sdt)[0];
  while (first && localName(first) === 'sdt') first = children(child(first, 'sdtContent') ?? first)[0];
  const name = first ? localName(first) : 'p';
  if (name === 'tc') return 'cell';
  if (name === 'tr') return 'row';
  if (name === 'r' || name === 'hyperlink' || name === 'fldSimple') return 'run';
  return localName(sdt.parentNode as Element) === 'p' ? 'run' : 'block';
}

export function textOf(el: Element): string {
  return Array.from(el.getElementsByTagNameNS(NS.w, 't'))
    .map((t) => t.textContent ?? '')
    .join('');
}

export function contentWidth(doc: Document): number {
  const sectPr = Array.from(doc.getElementsByTagNameNS(NS.w, 'sectPr')).pop();
  const pgSz = sectPr ? children(sectPr, 'pgSz')[0] : undefined;
  const pgMar = sectPr ? children(sectPr, 'pgMar')[0] : undefined;
  const n = (el: Element | undefined, a: string, d: number) => Number(el?.getAttributeNS(NS.w, a) || d);
  return n(pgSz, 'w', 11906) - n(pgMar, 'left', 1134) - n(pgMar, 'right', 1134);
}
