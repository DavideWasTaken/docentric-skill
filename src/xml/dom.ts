import { DOMParser, type Element } from '@xmldom/xmldom';

export type { Element };

/**
 * Docentric mixes default and prefixed namespaces freely (ProjectData has a
 * default namespace, its children reset it to "", attributes are `d2p1:`-prefixed
 * in one file and `ns1:` in the next). Everything here matches on localName only.
 */

export function parseXml(xml: string, partName: string): Element {
  const doc = new DOMParser({
    onError: (level, message) => {
      if (level !== 'warning') throw new Error(`Malformed XML in ${partName}: ${message}`);
    },
  }).parseFromString(xml, 'text/xml');
  const root = doc.documentElement;
  if (!root) throw new Error(`Empty XML part: ${partName}`);
  return root;
}

export function localName(el: Element): string {
  return el.localName ?? el.nodeName.split(':').pop() ?? el.nodeName;
}

export function children(el: Element, name?: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (name === undefined || localName(n as Element) === name)) out.push(n as Element);
  }
  return out;
}

export function child(el: Element, name: string): Element | null {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && localName(n as Element) === name) return n as Element;
  }
  return null;
}

/** Follows a chain of child local names, e.g. path(ds, 'Schema', 'XsdList', 'MainXsd'). */
export function path(el: Element, ...names: string[]): Element | null {
  let cur: Element | null = el;
  for (const name of names) {
    if (!cur) return null;
    cur = child(cur, name);
  }
  return cur;
}

export function descendants(el: Element, name: string): Element[] {
  const out: Element[] = [];
  const walk = (e: Element) => {
    for (const c of children(e)) {
      if (localName(c) === name) out.push(c);
      walk(c);
    }
  };
  walk(el);
  return out;
}

/** Attribute value by local name, whatever its prefix. */
export function attr(el: Element, name: string): string | null {
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (!a) continue;
    const ln = a.localName ?? a.name.split(':').pop();
    if (ln === name && !a.name.startsWith('xmlns')) return a.value;
  }
  return null;
}

export function allAttrs(el: Element): Record<string, string> {
  const out: Record<string, string> = {};
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (!a || a.name.startsWith('xmlns')) continue;
    out[a.localName ?? a.name] = a.value;
  }
  return out;
}
