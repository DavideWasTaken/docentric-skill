import { attr, child, children, type Element, localName } from '../xml/dom.js';
import type { ContentControl } from './types.js';

/** Docentric tags its content controls `<projectId>-<elementId>`, both decimal. */
const DOCENTRIC_TAG = /^(\d+)-(\d+)$/;

/** Parts that can carry template content. The glossary only holds placeholder building blocks. */
export function isContentPart(name: string): boolean {
  return /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name);
}

export function extractContentControls(root: Element, part: string): ContentControl[] {
  const out: ContentControl[] = [];
  const walk = (el: Element, parentElementId: string | null) => {
    for (const c of children(el)) {
      if (localName(c) !== 'sdt') {
        walk(c, parentElementId);
        continue;
      }
      const pr = child(c, 'sdtPr');
      const tagEl = pr ? child(pr, 'tag') : null;
      const aliasEl = pr ? child(pr, 'alias') : null;
      const tag = tagEl ? attr(tagEl, 'val') : null;
      const m = tag ? DOCENTRIC_TAG.exec(tag) : null;
      let nextParent = parentElementId;
      if (tag && m) {
        out.push({
          part,
          tag,
          projectId: m[1] ?? null,
          elementId: m[2] ?? null,
          alias: aliasEl ? attr(aliasEl, 'val') : null,
          parentElementId,
        });
        nextParent = m[2] ?? parentElementId;
      }
      const content = child(c, 'sdtContent');
      if (content) walk(content, nextParent);
    }
  };
  walk(root, null);
  return out;
}
