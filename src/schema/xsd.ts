import type { SchemaNode } from '../model/types.js';
import { attr, children, type Element, localName } from '../xml/dom.js';

/**
 * Turns the XSD that Docentric copies from the .ddsp into a tree of elements and
 * attributes. Only the subset Docentric emits is supported: named and inline
 * complexTypes, sequence/choice/all, attributes, element refs, and
 * complexContent/simpleContent extensions. Anything else is ignored rather than
 * rejected — a missing branch only means fewer paths can be resolved.
 */

const MAX_DEPTH = 40;

const stripPrefix = (qname: string) => qname.split(':').pop() ?? qname;

export function buildSchemaTree(schema: Element): SchemaNode | null {
  const types = new Map<string, Element>();
  const topElements = new Map<string, Element>();
  for (const c of children(schema)) {
    const name = attr(c, 'name');
    if (!name) continue;
    if (localName(c) === 'complexType') types.set(name, c);
    else if (localName(c) === 'element') topElements.set(name, c);
  }
  const first = topElements.values().next();
  if (first.done) return null;

  const buildElement = (el: Element, typeStack: string[]): SchemaNode => {
    const ref = attr(el, 'ref');
    const target = ref ? (topElements.get(stripPrefix(ref)) ?? el) : el;
    const node: SchemaNode = {
      name: attr(target, 'name') ?? stripPrefix(ref ?? '?'),
      attributes: new Map(),
      children: new Map(),
      repeated: isRepeated(el) || isRepeated(el.parentNode as Element | null),
    };
    if (typeStack.length > MAX_DEPTH) return node;

    const typeName = attr(target, 'type');
    if (typeName) {
      const named = types.get(stripPrefix(typeName));
      if (named && !typeStack.includes(typeName)) fillFromType(node, named, [...typeStack, typeName]);
    } else {
      for (const ct of children(target, 'complexType')) fillFromType(node, ct, typeStack);
    }
    return node;
  };

  const fillFromType = (node: SchemaNode, type: Element, typeStack: string[]) => {
    for (const c of children(type)) {
      switch (localName(c)) {
        case 'attribute': {
          const name = attr(c, 'name') ?? (attr(c, 'ref') ? stripPrefix(attr(c, 'ref') as string) : null);
          if (name) node.attributes.set(name, stripPrefix(attr(c, 'type') ?? 'string'));
          break;
        }
        case 'element': {
          const childNode = buildElement(c, typeStack);
          if (!node.children.has(childNode.name)) node.children.set(childNode.name, childNode);
          break;
        }
        case 'sequence':
        case 'choice':
        case 'all':
          fillFromType(node, c, typeStack);
          break;
        case 'complexContent':
        case 'simpleContent':
          for (const ext of children(c)) {
            const base = attr(ext, 'base');
            const baseType = base ? types.get(stripPrefix(base)) : undefined;
            if (baseType && base && !typeStack.includes(base)) fillFromType(node, baseType, [...typeStack, base]);
            fillFromType(node, ext, typeStack);
          }
          break;
        default:
          break;
      }
    }
  };

  return buildElement(first.value, []);
}

function isRepeated(el: Element | null): boolean {
  if (el?.nodeType !== 1) return false;
  const max = attr(el, 'maxOccurs');
  return max === 'unbounded' || (max !== null && Number(max) > 1);
}

export function countSchema(node: SchemaNode): { elements: number; attributes: number } {
  let elements = 1;
  let attributes = node.attributes.size;
  for (const c of node.children.values()) {
    const sub = countSchema(c);
    elements += sub.elements;
    attributes += sub.attributes;
  }
  return { elements, attributes };
}
