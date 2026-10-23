import { resolve } from 'node:path';
import { type Document, type Element, type Node, XMLSerializer } from '@xmldom/xmldom';
import { z } from 'zod';
import { IMAGE_REL_TOKEN, LayoutCompiler, type Scope } from '../create/compile.js';
import { BlockSchema } from '../create/spec.js';
import { contentParts, findProjectDataPart, loadTemplate } from '../docx/loadTemplate.js';
import { NS, REL, readTemplateBytes, TemplateError, WordPackage } from '../docx/package.js';
import { TEMPLATE_LABEL_PREFIX } from '../model/projectData.js';
import type { DocentricTemplate } from '../model/types.js';
import { type Finding, type ValidationResult, validateTemplate } from '../validate/validateTemplate.js';
import { solidPng } from '../write/ooxml.js';
import { IdAllocator, type NewElement } from '../write/projectDataXml.js';
import { child, children, localName } from '../xml/dom.js';
import {
  ancestor,
  contentWidth,
  enclosingElementId,
  findSdt,
  insertNodes,
  isInside,
  parseWordFragment,
  Registry,
  sdtLevel,
  sdtTag,
  textOf,
} from './dom.js';

/**
 * edit: a list of operations applied to a COPY of a template. After the
 * operations the result is validated; if it has errors the original did not
 * have, nothing is written (unless force). Reference: docs/EDIT_OPERATIONS.md.
 */

const ElementRef = z
  .string()
  .regex(/^#?\d+$/, 'element id, e.g. "#1200000001" (from the inspect command)')
  .transform((s) => s.replace('#', ''));

const Anchor = z.union([z.object({ element: ElementRef }).strict(), z.object({ text: z.string().min(1) }).strict()]);

export const EditOp = z.discriminatedUnion('op', [
  z
    .object({
      op: z.literal('insert'),
      position: z.enum(['after', 'before', 'inside']).default('after'),
      anchor: Anchor,
      blocks: z.array(BlockSchema).min(1),
    })
    .strict(),
  z
    .object({
      op: z.literal('add_column'),
      list: ElementRef,
      after: ElementRef.optional(),
      header: z.string().optional(),
      value: z.string().min(1),
      width: z.number().positive().optional(),
      widthCm: z.number().min(0.3).max(30).optional(),
      align: z.enum(['left', 'center', 'right']).optional(),
      format: z.string().optional(),
    })
    .strict(),
  z
    .object({
      op: z.literal('set_binding'),
      element: ElementRef,
      path: z.string().min(1).optional(),
      source: z.string().optional(),
      format: z.string().nullable().optional(),
    })
    .strict(),
  z
    .object({
      op: z.literal('wrap_if'),
      element: ElementRef,
      condition: z.string().min(1),
      source: z.string().optional(),
      level: z.enum(['auto', 'paragraph', 'row']).default('auto'),
    })
    .strict(),
  z.object({ op: z.literal('remove'), element: ElementRef, keepContent: z.boolean().default(false) }).strict(),
  z.object({ op: z.literal('set_label'), code: z.string().min(1), texts: z.record(z.string(), z.string()) }).strict(),
  z.object({ op: z.literal('convert_placeholders') }).strict(),
  z
    .object({
      op: z.literal('resize_column'),
      element: ElementRef,
      widthCm: z.number().min(0.3).max(30),
    })
    .strict(),
  z.object({ op: z.literal('remove_column'), element: ElementRef }).strict(),
  z
    .object({
      op: z.literal('set_style'),
      anchor: Anchor,
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      caps: z.boolean().optional(),
      size: z.number().min(4).max(96).optional(),
      color: z
        .string()
        .regex(/^#?[0-9A-Fa-f]{6}$/)
        .optional(),
      align: z.enum(['left', 'center', 'right', 'justify']).optional(),
    })
    .strict(),
  z
    .object({
      op: z.literal('add_language'),
      language: z.string().regex(/^[a-z]{2,3}(-[a-z0-9]{2,8})*$/i, 'a language code such as de, it, en-us'),
      default: z.boolean().default(false),
    })
    .strict(),
]);
export type EditOp = z.infer<typeof EditOp>;

export interface AppliedOp {
  op: string;
  detail: string;
  newElements: string[];
}

export interface EditResult {
  written: boolean;
  path: string | null;
  applied: AppliedOp[];
  before: ValidationResult;
  after: ValidationResult;
  /** Errors present after the edit and not before. */
  newErrors: Finding[];
}

export interface EditOptions {
  outputPath?: string;
  dryRun?: boolean;
  /** Write even if the edit introduces validation errors. */
  force?: boolean;
}

export class EditError extends Error {}

export function parseOps(input: unknown): EditOp[] {
  const r = z.array(EditOp).min(1).safeParse(input);
  if (!r.success) {
    const msg = r.error.issues
      .slice(0, 10)
      .map((i) => `- ops.${i.path.join('.')}: ${i.message}`)
      .join('\n');
    throw new EditError(`Invalid operations:\n${msg}`);
  }
  return r.data;
}

export async function editTemplate(
  templatePath: string,
  opsInput: unknown,
  opts: EditOptions = {},
): Promise<EditResult> {
  const ops = parseOps(opsInput);
  const source = resolve(templatePath);
  const output = resolve(opts.outputPath ?? source.replace(/(\.\w+)$/, '.edited$1'));
  if (output.toLowerCase() === source.toLowerCase()) {
    throw new EditError('outputPath must differ from the template: edits always go to a new file.');
  }

  const bytes = await readTemplateBytes(source);
  const before = validateTemplate(await loadTemplate(bytes));
  const session = await EditSession.open(bytes);
  const applied: AppliedOp[] = [];
  for (const [i, op] of ops.entries()) {
    try {
      applied.push(await session.apply(op));
    } catch (e) {
      throw new EditError(`Operation ${i + 1} (${op.op}) failed: ${(e as Error).message}`);
    }
  }
  const pkg = await session.commit();
  const after = validateTemplate(await loadTemplate(pkg));

  const key = (f: Finding) => `${f.code}|${f.message}`;
  const known = new Set(before.findings.filter((f) => f.severity === 'error').map(key));
  const newErrors = after.findings.filter((f) => f.severity === 'error' && !known.has(key(f)));

  const write = !opts.dryRun && (newErrors.length === 0 || !!opts.force);
  if (write) await pkg.saveAs(output);
  return { written: write, path: write ? output : null, applied, before, after, newErrors };
}

class EditSession {
  private readonly parts = new Map<string, Document>();
  private ids!: IdAllocator;
  private projectId!: string;
  private template!: DocentricTemplate;
  private registry!: Registry;
  private pdPart!: string;
  private defaultSource = 'MainData';

  private constructor(private readonly pkg: WordPackage) {}

  static async open(bytes: Buffer): Promise<EditSession> {
    const s = new EditSession(await WordPackage.open(bytes));
    s.template = await loadTemplate(s.pkg);
    s.pdPart = (await findProjectDataPart(s.pkg)) as string;
    s.registry = new Registry((await s.pkg.dom(s.pdPart)) as Document);
    s.ids = new IdAllocator(s.registry.ids());
    s.projectId = s.template.contentControls[0]?.projectId ?? IdAllocator.projectId();
    const names = s.template.dataSources.map((d) => d.name);
    s.defaultSource = names.includes('MainData') ? 'MainData' : (names[0] ?? 'MainData');
    for (const part of contentParts(s.pkg)) {
      const doc = await s.pkg.dom(part);
      if (doc) s.parts.set(part, doc);
    }
    return s;
  }

  async commit(): Promise<WordPackage> {
    for (const [part, doc] of this.parts) this.pkg.setDom(part, doc);
    this.pkg.setDom(this.pdPart, this.registry.doc);
    return this.pkg;
  }

  apply(op: EditOp): Promise<AppliedOp> | AppliedOp {
    switch (op.op) {
      case 'insert':
        return this.insert(op);
      case 'add_column':
        return this.addColumn(op);
      case 'set_binding':
        return this.setBinding(op);
      case 'wrap_if':
        return this.wrapIf(op);
      case 'remove':
        return this.remove(op);
      case 'set_label':
        return this.setLabel(op);
      case 'convert_placeholders':
        return this.convertPlaceholders();
      case 'add_language':
        return this.addLanguage(op);
      case 'resize_column':
        return this.resizeColumn(op);
      case 'remove_column':
        return this.removeColumn(op);
      case 'set_style':
        return this.setStyle(op);
    }
  }

  // --- helpers -------------------------------------------------------------------

  private locate(elementId: string): { part: string; doc: Document; sdt: Element } {
    for (const [part, doc] of this.parts) {
      const sdt = findSdt(doc, elementId);
      if (sdt) return { part, doc, sdt };
    }
    if (this.registry.element(elementId))
      throw new EditError(`element ${elementId} is registered but not placed in the document`);
    throw new EditError(`no element ${elementId}; run inspect to list element ids`);
  }

  private typeOf(id: string | null): string | null {
    return id ? (this.registry.element(id)?.getAttribute('Type') ?? null) : null;
  }

  /** Source for new bindings under `parentId`: '.' inside a List/Group, else the default data source. */
  private scopeSource(parentId: string | null): string {
    for (let p = parentId; p; p = this.registry.parentOf(p)) {
      const t = this.typeOf(p);
      if (t === 'List' || t === 'Group') return '.';
    }
    return this.defaultSource;
  }

  private compiler(): LayoutCompiler {
    return new LayoutCompiler(
      this.projectId,
      this.ids,
      this.template.dataSources.map((d) => d.name),
    );
  }

  /** Adds compiled elements to the registry; top-level ones go under `parentId` after `afterId`. */
  private registerCompiled(compiler: LayoutCompiler, parentId: string | null, afterId?: string | null): string[] {
    const nested = new Set(compiler.elements.flatMap((e) => e.children));
    let after = afterId ?? null;
    for (const e of compiler.elements) {
      this.registry.add(e);
      if (parentId && !nested.has(e.id)) {
        this.registry.addChild(parentId, e.id, after);
        after = e.id;
      }
    }
    return compiler.elements.map((e) => e.id);
  }

  private async resolveImages(part: string, xml: string): Promise<string> {
    if (!xml.includes(IMAGE_REL_TOKEN)) return xml;
    const media = 'word/media/docentric-placeholder.png';
    if (!this.pkg.has(media)) this.pkg.setBytes(media, solidPng(16, 8, [217, 217, 217]));
    await this.pkg.ensureDefaultContentType('png', 'image/png');
    const relId = await this.pkg.addRelationship(part, REL.image, 'media/docentric-placeholder.png');
    return xml.split(IMAGE_REL_TOKEN).join(relId);
  }

  // --- insert --------------------------------------------------------------------

  private async insert(op: Extract<EditOp, { op: 'insert' }>): Promise<AppliedOp> {
    let part: string;
    let doc: Document;
    let container: Node;
    let before: Node | null;
    let parentId: string | null;
    let afterId: string | null = null;
    let width: number;

    if ('element' in op.anchor) {
      const loc = this.locate(op.anchor.element);
      part = loc.part;
      doc = loc.doc;
      width = contentWidth(doc);
      const level = sdtLevel(loc.sdt);
      if (op.position === 'inside') {
        const type = this.typeOf(op.anchor.element);
        if (!['If', 'List', 'Group'].includes(type ?? ''))
          throw new EditError(`'inside' needs an If/List/Group, not a ${type}`);
        if (level === 'row')
          throw new EditError('this List repeats table rows: use add_column, or insert after the table');
        container = child(loc.sdt, 'sdtContent') as Element;
        before = null;
        parentId = op.anchor.element;
        afterId = this.registry.childIds(parentId).pop() ?? null;
      } else {
        parentId = this.registry.parentOf(op.anchor.element);
        afterId = op.anchor.element;
        let ref: Element;
        if (level === 'run') ref = ancestor(loc.sdt, 'p') as Element;
        else if (level === 'row') ref = ancestor(loc.sdt, 'tbl') as Element;
        else if (level === 'cell') {
          // "after a field in a table cell" means: inside that cell, below its content.
          const tc = children(child(loc.sdt, 'sdtContent') as Element, 'tc')[0] as Element;
          const paras = children(tc).filter((c) => localName(c) !== 'tcPr');
          ref = (op.position === 'after' ? paras[paras.length - 1] : paras[0]) as Element;
          width = Number(child(child(tc, 'tcPr') ?? tc, 'tcW')?.getAttributeNS(NS.w, 'w') || width / 2);
        } else ref = loc.sdt;
        container = ref.parentNode as Node;
        before = op.position === 'after' ? ref.nextSibling : ref;
      }
    } else {
      const found = this.findParagraph(op.anchor.text);
      if (!found) throw new EditError(`no paragraph contains "${op.anchor.text}"`);
      ({ part, doc } = found);
      width = contentWidth(doc);
      if (op.position === 'inside') throw new EditError("'inside' needs an element anchor");
      parentId = enclosingElementId(found.p);
      container = found.p.parentNode as Node;
      before = op.position === 'after' ? found.p.nextSibling : found.p;
    }

    const compiler = this.compiler();
    const scope: Scope = { source: this.scopeSource(parentId), parentId, width };
    const xml = await this.resolveImages(part, compiler.blocks(op.blocks, scope));
    insertNodes(doc, parseWordFragment(xml), container, before);
    const created = this.registerCompiled(compiler, parentId, afterId);
    return {
      op: 'insert',
      detail: `${op.blocks.length} block(s) ${op.position} ${describeAnchor(op.anchor)} in ${part}`,
      newElements: created,
    };
  }

  private findParagraph(text: string): { part: string; doc: Document; p: Element } | null {
    for (const [part, doc] of this.parts) {
      for (const p of Array.from(doc.getElementsByTagNameNS(NS.w, 'p'))) {
        if (textOf(p).includes(text)) return { part, doc, p };
      }
    }
    return null;
  }

  // --- add_column ----------------------------------------------------------------

  private addColumn(op: Extract<EditOp, { op: 'add_column' }>): AppliedOp {
    if (this.typeOf(op.list) !== 'List') throw new EditError(`element ${op.list} is not a List`);
    const loc = this.locate(op.list);
    if (sdtLevel(loc.sdt) !== 'row') throw new EditError(`List ${op.list} does not repeat table rows`);
    const tbl = ancestor(loc.sdt, 'tbl') as Element;
    const rows = tableRows(tbl);
    const listRows = rows.filter((tr) => isInside(tr, loc.sdt));
    // The row that carries the record's fields: the widest one (others are notes rows spanning the table).
    const dataRow = listRows.reduce(
      (a, b) => (cellUnits(b).length > cellUnits(a).length ? b : a),
      listRows[0] as Element,
    );
    // Header: the nearest row above the List that has text (spacer rows are skipped).
    const headerRow =
      rows
        .filter((tr) => !isInside(tr, loc.sdt) && isBefore(tr, loc.sdt))
        .reverse()
        .find((tr) => textOf(tr).trim() !== '') ?? null;

    // Grid column where the new column goes.
    let gridIndex = rowGridWidth(dataRow);
    if (op.after) {
      const after = this.locate(op.after).sdt;
      const tc =
        sdtLevel(after) === 'cell'
          ? (children(child(after, 'sdtContent') as Element, 'tc')[0] as Element)
          : ancestor(after, 'tc');
      if (!tc || ancestor(tc, 'tr') !== dataRow) throw new EditError(`element ${op.after} is not in the List's row`);
      gridIndex = cellGridStart(dataRow, tc) + cellSpan(tc);
    }

    // Grid: insert the column; the table keeps its width.
    const grid = child(tbl, 'tblGrid') as Element;
    const cols = children(grid, 'gridCol');
    const widths = cols.map((c) => Number(c.getAttributeNS(NS.w, 'w') || 0));
    const wanted = op.widthCm ? Math.round(op.widthCm * 567) : newColumnWidth(dataRow, widths, op.width ?? 1);
    const newGrid = takeSpace(widths, gridIndex, wanted);
    cols.forEach((c, i) => {
      c.setAttributeNS(NS.w, 'w:w', String(newGrid[i < gridIndex ? i : i + 1]));
    });
    const gc = loc.doc.createElementNS(NS.w, 'w:gridCol');
    const colWidth = newGrid[gridIndex] as number;
    gc.setAttributeNS(NS.w, 'w:w', String(colWidth));
    grid.insertBefore(gc, cols[gridIndex] ?? null);

    const listParent = this.registry.parentOf(op.list);
    const created: string[] = [];
    for (const tr of rows) {
      const neighbour = cellAtGrid(tr, Math.max(0, gridIndex - 1)) ?? cellAtGrid(tr, gridIndex);
      // A cell spanning the insertion point just grows.
      const spanning = cellCovering(tr, gridIndex);
      if (spanning && cellGridStart(tr, spanning) < gridIndex) {
        setSpan(spanning, cellSpan(spanning) + 1);
        continue;
      }
      let xml: string;
      const compiler = this.compiler();
      if (tr === dataRow) {
        const scope: Scope = { source: '.', parentId: op.list, width: colWidth };
        xml = compiler.cellXml(op.value, scope, { width: colWidth, borders: true }, op.format);
        created.push(...this.registerCompiled(compiler, op.list, null));
      } else if (tr === headerRow && op.header) {
        const scope: Scope = { source: this.scopeSource(listParent), parentId: listParent, width: colWidth };
        xml = compiler.cellXml(op.header, scope, { width: colWidth, borders: true });
        created.push(...this.registerCompiled(compiler, listParent, op.list));
      } else {
        xml = `<w:tc><w:tcPr><w:tcW w:w="${colWidth}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
      }
      const [node] = insertNodes(loc.doc, parseWordFragment(xml), tr, cellUnitAtGrid(tr, gridIndex));
      if (node && neighbour) copyCellLook(node, neighbour, colWidth, op.align);
    }
    syncCellWidths(tbl);
    const shrunk = newGrid
      .filter((_, i) => i !== gridIndex)
      .map((w, i) => ({ i, before: widths[i] as number, after: w }))
      .filter((c) => c.after < c.before);
    return {
      op: 'add_column',
      detail:
        `column at position ${gridIndex + 1} of List ${op.list} (${op.value}), ${twipsToCm(colWidth)} cm wide; ` +
        `space taken from ${shrunk.length} wider column(s): ${shrunk.map((c) => `${twipsToCm(c.before)}→${twipsToCm(c.after)} cm`).join(', ')}`,
      newElements: created,
    };
  }

  // --- set_binding ----------------------------------------------------------------

  private setBinding(op: Extract<EditOp, { op: 'set_binding' }>): AppliedOp {
    const binding = this.registry.binding(op.element);
    if (!binding) throw new EditError(`element ${op.element} has no binding`);
    const changes: string[] = [];
    if (op.path !== undefined) {
      binding.setAttribute('Path', op.path);
      changes.push(`path=${op.path}`);
      const { sdt } = this.locate(op.element);
      if (this.typeOf(op.element) === 'Field') setFirstText(sdt, op.path);
    }
    if (op.source !== undefined) {
      binding.setAttribute('Source', op.source);
      changes.push(`source=${op.source}`);
    }
    if (op.format !== undefined) {
      this.registry.setAttribute(op.element, 'FormatString', op.format);
      changes.push(`format=${op.format ?? '(none)'}`);
    }
    return { op: 'set_binding', detail: `${op.element}: ${changes.join(', ') || 'no change'}`, newElements: [] };
  }

  // --- wrap_if ---------------------------------------------------------------------

  private wrapIf(op: Extract<EditOp, { op: 'wrap_if' }>): AppliedOp {
    const loc = this.locate(op.element);
    const level = sdtLevel(loc.sdt);
    let target: Element;
    if (op.level === 'row') target = ancestor(loc.sdt, 'tr') as Element;
    else if (level === 'run' || op.level === 'paragraph')
      target = (level === 'run' ? ancestor(loc.sdt, 'p') : loc.sdt) as Element;
    else target = loc.sdt; // block, cell or row level: wrap the content control itself
    if (!target) throw new EditError(`cannot find a ${op.level} around element ${op.element}`);

    const oldParent = this.registry.parentOf(op.element);
    const id = this.ids.elementId();
    const el: NewElement = {
      id,
      type: 'If',
      source: op.source ?? this.scopeSource(oldParent),
      path: op.condition,
      children: [],
    };
    this.registry.add(el);
    const wrapped = topLevelElementIds(target);
    if (oldParent) this.registry.addChild(oldParent, id, wrapped[0] ?? op.element);
    for (const w of wrapped) {
      const p = this.registry.parentOf(w);
      if (p) this.registry.removeChild(p, w);
      this.registry.addChild(id, w);
    }

    const doc = loc.doc;
    const sdt = doc.createElementNS(NS.w, 'w:sdt');
    const pr = parseWordFragment(
      `<w:sdtPr><w:alias w:val="If"/><w:tag w:val="${this.projectId}-${id}"/><w:id w:val="${IdAllocator.sdtId()}"/><w15:color w:val="0000FF"/></w:sdtPr>`,
    );
    insertNodes(doc, pr, sdt, null);
    const content = doc.createElementNS(NS.w, 'w:sdtContent');
    sdt.appendChild(content);
    target.parentNode?.insertBefore(sdt, target);
    content.appendChild(target);
    return { op: 'wrap_if', detail: `If ${id} (${op.condition}) around ${describeNode(target)}`, newElements: [id] };
  }

  // --- remove ----------------------------------------------------------------------

  private remove(op: Extract<EditOp, { op: 'remove' }>): AppliedOp {
    const loc = this.locate(op.element);
    const parent = loc.sdt.parentNode as Element;
    const content = child(loc.sdt, 'sdtContent');
    if (op.keepContent) {
      const parentId = this.registry.parentOf(op.element);
      for (const c of this.registry.childIds(op.element)) if (parentId) this.registry.addChild(parentId, c, op.element);
      for (const c of content ? children(content) : []) parent.insertBefore(c, loc.sdt);
      parent.removeChild(loc.sdt);
      const el = this.registry.element(op.element);
      if (parentId) this.registry.removeChild(parentId, op.element);
      el?.parentNode?.removeChild(el);
      return { op: 'remove', detail: `element ${op.element} removed, content kept`, newElements: [] };
    }
    const inner = Array.from(loc.sdt.getElementsByTagNameNS(NS.w, 'sdt'))
      .map((s) => sdtTag(s)?.elementId)
      .filter(Boolean) as string[];
    const level = sdtLevel(loc.sdt);
    if (level === 'cell') {
      // Keep the grid intact: the cell stays, empty.
      const tc = children(content as Element, 'tc')[0] as Element;
      for (const c of children(tc)) if (localName(c) !== 'tcPr') tc.removeChild(c);
      tc.appendChild(loc.doc.createElementNS(NS.w, 'w:p'));
      parent.insertBefore(tc, loc.sdt);
    }
    parent.removeChild(loc.sdt);
    const removed = new Set(this.registry.remove(op.element));
    for (const id of inner) if (!removed.has(id)) for (const r of this.registry.remove(id)) removed.add(r);
    return { op: 'remove', detail: `removed ${removed.size} element(s): ${[...removed].join(', ')}`, newElements: [] };
  }

  // --- labels ------------------------------------------------------------------------

  private setLabel(op: Extract<EditOp, { op: 'set_label' }>): AppliedOp {
    const code = op.code.startsWith('@') ? op.code : `${TEMPLATE_LABEL_PREFIX}${op.code}`;
    const result = this.registry.upsertLabel({ code, texts: op.texts, custom: true });
    const languages = this.registry.languages();
    const missing = languages.filter((l) => !(l in op.texts));
    return {
      op: 'set_label',
      detail: `${code} ${result}${missing.length && result === 'added' ? ` (no text yet for: ${missing.join(', ')})` : ''}`,
      newElements: [],
    };
  }

  // --- resize_column -----------------------------------------------------------------

  /**
   * Sets the width of the table column that holds `element`. The table keeps its width:
   * extra space comes from the other columns wider than average, freed space goes to the
   * widest other column.
   */
  private resizeColumn(op: Extract<EditOp, { op: 'resize_column' }>): AppliedOp {
    const { sdt } = this.locate(op.element);
    const tc =
      sdtLevel(sdt) === 'cell'
        ? (children(child(sdt, 'sdtContent') as Element, 'tc')[0] as Element)
        : ancestor(sdt, 'tc');
    if (!tc) throw new EditError(`element ${op.element} is not in a table cell`);
    const tr = ancestor(tc, 'tr') as Element;
    const tbl = ancestor(tr, 'tbl') as Element;
    const start = cellGridStart(tr, tc);
    const span = cellSpan(tc);
    const cols = children(child(tbl, 'tblGrid') as Element, 'gridCol');
    const widths = cols.map((c) => Number(c.getAttributeNS(NS.w, 'w') || 0));
    const resized = resizeSpan(widths, start, span, Math.round(op.widthCm * 567));
    cols.forEach((c, i) => {
      c.setAttributeNS(NS.w, 'w:w', String(resized[i]));
    });
    syncCellWidths(tbl);
    const changed = resized
      .map((w, i) => ({ i, before: widths[i] as number, after: w }))
      .filter((c) => c.after !== c.before && (c.i < start || c.i >= start + span));
    return {
      op: 'resize_column',
      detail:
        `column of #${op.element}: ${twipsToCm(widths.slice(start, start + span).reduce((a, b) => a + b, 0))} → ${op.widthCm} cm; ` +
        `other columns: ${changed.map((c) => `${twipsToCm(c.before)}→${twipsToCm(c.after)} cm`).join(', ') || 'unchanged'}`,
      newElements: [],
    };
  }

  // --- remove_column -----------------------------------------------------------------

  /**
   * Removes the table column holding `element`: its cell in every row (with every Docentric
   * element inside, from the registry too), its grid columns; cells spanning across it shrink.
   * The freed width is spread over the remaining columns, so the table keeps its width.
   */
  private removeColumn(op: Extract<EditOp, { op: 'remove_column' }>): AppliedOp {
    const { sdt } = this.locate(op.element);
    const tc =
      sdtLevel(sdt) === 'cell'
        ? (children(child(sdt, 'sdtContent') as Element, 'tc')[0] as Element)
        : ancestor(sdt, 'tc');
    if (!tc) throw new EditError(`element ${op.element} is not in a table cell`);
    const tbl = ancestor(ancestor(tc, 'tr') as Element, 'tbl') as Element;
    const start = cellGridStart(ancestor(tc, 'tr') as Element, tc);
    const span = cellSpan(tc);
    const end = start + span;

    const removedIds = new Set<string>();
    let cells = 0;
    for (const tr of tableRows(tbl)) {
      let col = 0;
      for (const { unit, tc: cell } of cellUnits(tr)) {
        const s = col;
        const n = cellSpan(cell);
        col += n;
        const overlap = Math.min(s + n, end) - Math.max(s, start);
        if (overlap <= 0) continue;
        if (s >= start && s + n <= end) {
          // The cell lies inside the column: it goes, with its elements.
          for (const id of [...topLevelElementIds(unit)]) for (const r of this.registry.remove(id)) removedIds.add(r);
          unit.parentNode?.removeChild(unit);
          cells++;
        } else {
          setSpan(cell, n - overlap);
        }
      }
    }

    const grid = child(tbl, 'tblGrid') as Element;
    const cols = children(grid, 'gridCol');
    const freed = cols.slice(start, end).reduce((n, c) => n + Number(c.getAttributeNS(NS.w, 'w') || 0), 0);
    for (const c of cols.slice(start, end)) grid.removeChild(c);
    // The freed width is spread over the remaining columns in proportion to their width, as Word does.
    const rest = children(grid, 'gridCol');
    const restWidths = rest.map((c) => Number(c.getAttributeNS(NS.w, 'w') || 0));
    const restTotal = restWidths.reduce((a, b) => a + b, 0) || 1;
    let given = 0;
    rest.forEach((c, i) => {
      const extra = i === rest.length - 1 ? freed - given : Math.round(((restWidths[i] as number) / restTotal) * freed);
      given += extra;
      c.setAttributeNS(NS.w, 'w:w', String((restWidths[i] as number) + extra));
    });
    syncCellWidths(tbl);
    return {
      op: 'remove_column',
      detail: `removed ${cells} cell(s) and ${removedIds.size} element(s) (${[...removedIds].map((i) => `#${i}`).join(', ')}); ${twipsToCm(freed)} cm spread over the other columns`,
      newElements: [],
    };
  }

  // --- set_style ---------------------------------------------------------------------

  /** Formats the runs of an element (or of a paragraph found by its text): bold, size, colour, alignment… */
  private setStyle(op: Extract<EditOp, { op: 'set_style' }>): AppliedOp {
    let target: Element;
    if ('element' in op.anchor) {
      const { sdt } = this.locate(op.anchor.element);
      target = sdtLevel(sdt) === 'run' && op.align ? (ancestor(sdt, 'p') as Element) : sdt;
    } else {
      const found = this.findParagraph(op.anchor.text);
      if (!found) throw new EditError(`no paragraph contains "${op.anchor.text}"`);
      target = found.p;
    }
    const doc = target.ownerDocument as Document;
    const runs = Array.from(target.getElementsByTagNameNS(NS.w, 'r'));
    // The content control's own rPr formats the text the Designer shows while editing.
    const sdtPrs = localName(target) === 'sdt' ? [child(target, 'sdtPr') as Element] : [];
    const applied: string[] = [];
    const changes: Array<[string, Record<string, string> | undefined]> = [
      ['b', op.bold === undefined ? undefined : op.bold ? {} : { 'w:val': '0' }],
      ['bCs', op.bold === undefined ? undefined : op.bold ? {} : { 'w:val': '0' }],
      ['i', op.italic === undefined ? undefined : op.italic ? {} : { 'w:val': '0' }],
      ['iCs', op.italic === undefined ? undefined : op.italic ? {} : { 'w:val': '0' }],
      ['caps', op.caps === undefined ? undefined : op.caps ? {} : { 'w:val': '0' }],
      ['color', op.color ? { 'w:val': op.color.replace('#', '').toUpperCase() } : undefined],
      ['sz', op.size ? { 'w:val': String(Math.round(op.size * 2)) } : undefined],
      ['szCs', op.size ? { 'w:val': String(Math.round(op.size * 2)) } : undefined],
      ['u', op.underline === undefined ? undefined : { 'w:val': op.underline ? 'single' : 'none' }],
    ];
    for (const r of runs) {
      let rPr = child(r, 'rPr');
      if (!rPr) {
        rPr = doc.createElementNS(NS.w, 'w:rPr');
        r.insertBefore(rPr, r.firstChild);
      }
      for (const [name, attrs] of changes) if (attrs !== undefined) setOrdered(rPr, name, attrs, RPR_ORDER);
    }
    for (const pr of sdtPrs) {
      let rPr = child(pr, 'rPr');
      if (!rPr) {
        rPr = doc.createElementNS(NS.w, 'w:rPr');
        pr.insertBefore(rPr, pr.firstChild);
      }
      for (const [name, attrs] of changes) if (attrs !== undefined) setOrdered(rPr, name, attrs, RPR_ORDER);
    }
    for (const [name, attrs] of changes) if (attrs !== undefined && !name.endsWith('Cs')) applied.push(name);
    if (op.align) {
      const paragraphs = localName(target) === 'p' ? [target] : Array.from(target.getElementsByTagNameNS(NS.w, 'p'));
      for (const p of paragraphs) {
        let pPr = child(p, 'pPr');
        if (!pPr) {
          pPr = doc.createElementNS(NS.w, 'w:pPr');
          p.insertBefore(pPr, p.firstChild);
        }
        setOrdered(pPr, 'jc', { 'w:val': op.align === 'justify' ? 'both' : op.align }, PPR_ORDER);
      }
      applied.push(`align=${op.align}`);
    }
    if (!applied.length)
      throw new EditError('set_style needs at least one of bold, italic, underline, caps, size, color, align');
    return {
      op: 'set_style',
      detail: `${applied.join(', ')} on ${runs.length} run(s) of ${'element' in op.anchor ? `#${op.anchor.element}` : `"${op.anchor.text}"`}`,
      newElements: [],
    };
  }

  // --- languages ---------------------------------------------------------------------

  private addLanguage(op: Extract<EditOp, { op: 'add_language' }>): AppliedOp {
    const language = op.language.toLowerCase();
    const added = this.registry.addLanguage(language, op.default);
    return {
      op: 'add_language',
      detail:
        `${language} ${added ? 'added' : 'already present'}${op.default ? ', now the default' : ''}` +
        ' — add its texts with set_label (validate lists the labels still missing it as DCX009)',
      newElements: [],
    };
  }

  // --- placeholders ------------------------------------------------------------------

  private convertPlaceholders(): AppliedOp {
    const created: string[] = [];
    const skipped: string[] = [];
    let converted = 0;
    const serializer = new XMLSerializer();
    for (const [part, doc] of this.parts) {
      for (const p of Array.from(doc.getElementsByTagNameNS(NS.w, 'p'))) {
        const text = textOf(p);
        if (!/\{\{[\s\S]+?\}\}/.test(text)) continue;
        const plain = children(p).every(
          (c) =>
            ['pPr', 'r', 'proofErr', 'bookmarkStart', 'bookmarkEnd'].includes(localName(c)) &&
            (localName(c) !== 'r' ||
              children(c).every((x) => ['rPr', 't', 'tab', 'br', 'lastRenderedPageBreak'].includes(localName(x)))),
        );
        if (!plain) {
          skipped.push(`"${text.slice(0, 40)}" in ${part} (has fields, links or drawings)`);
          continue;
        }
        const parentId = enclosingElementId(p);
        const firstRun = children(p, 'r')[0];
        const rPr = firstRun ? child(firstRun, 'rPr') : null;
        const compiler = this.compiler();
        const scope: Scope = { source: this.scopeSource(parentId), parentId, width: contentWidth(doc) };
        let xml: string;
        try {
          // Serialising a lone node repeats the namespace declaration; the fragment declares it already.
          const raw = rPr ? serializer.serializeToString(rPr).replace(/\s+xmlns:\w+="[^"]*"/g, '') : '';
          xml = compiler.inlineWithRawRunProps(text, scope, raw);
        } catch (e) {
          skipped.push(`"${text.slice(0, 40)}": ${(e as Error).message}`);
          continue;
        }
        for (const c of children(p)) if (localName(c) !== 'pPr') p.removeChild(c);
        insertNodes(doc, parseWordFragment(xml), p, null);
        created.push(...this.registerCompiled(compiler, parentId, null));
        converted += (text.match(/\{\{[\s\S]+?\}\}/g) ?? []).length;
      }
    }
    if (converted === 0 && skipped.length === 0) throw new EditError('no {{…}} placeholders found in the document');
    return {
      op: 'convert_placeholders',
      detail: `${converted} placeholder(s) converted${skipped.length ? `; skipped: ${skipped.join('; ')}` : ''}`,
      newElements: created,
    };
  }
}

// --- table helpers ------------------------------------------------------------------

/** Rows of a table, including rows wrapped in content controls, excluding nested tables. */
function tableRows(tbl: Element): Element[] {
  return Array.from(tbl.getElementsByTagNameNS(NS.w, 'tr')).filter((tr) => ancestor(tr, 'tbl') === tbl);
}

const twipsToCm = (t: number) => (t / 567).toFixed(1);

/**
 * Width of a new column: `weight` × the average width of the cells in the data row
 * (not of grid columns, which can be tiny spacers in Designer-made tables).
 */
function newColumnWidth(dataRow: Element, grid: number[], weight: number): number {
  const units = cellUnits(dataRow);
  let col = 0;
  const cellWidths = units.map((u) => {
    const span = cellSpan(u.tc);
    const w = grid.slice(col, col + span).reduce((a, b) => a + b, 0);
    col += span;
    return w;
  });
  const avg = cellWidths.reduce((a, b) => a + b, 0) / Math.max(1, cellWidths.length);
  return Math.max(400, Math.round(avg * weight));
}

/** Narrowest width a column is trimmed to when making room (~0.7 cm). */
const MIN_COLUMN = 400;

/**
 * Frees `amount` twips by trimming the widest eligible columns first, down to a common
 * level — what a person does: the wide Description column gives way, and the next widest
 * only once they are equal. Narrow columns are untouched. If even trimming everything to
 * MIN_COLUMN is not enough, all eligible columns shrink in proportion.
 */
export function shaveFromTop(widths: number[], amount: number, eligible: (i: number) => boolean): number[] {
  const idx = widths.map((_, i) => i).filter(eligible);
  const freedAt = (level: number) => idx.reduce((n, i) => n + Math.max(0, (widths[i] as number) - level), 0);
  if (amount <= 0 || idx.length === 0) return widths.slice();
  if (freedAt(MIN_COLUMN) < amount) {
    const pool = idx.reduce((n, i) => n + (widths[i] as number), 0);
    return widths.map((w, i) => (eligible(i) ? w * (1 - amount / pool) : w));
  }
  let lo = MIN_COLUMN;
  let hi = Math.max(...idx.map((i) => widths[i] as number));
  for (let k = 0; k < 40; k++) {
    const mid = (lo + hi) / 2;
    if (freedAt(mid) > amount) lo = mid;
    else hi = mid;
  }
  return widths.map((w, i) => (eligible(i) ? Math.min(w, hi) : w));
}

/**
 * Returns the grid with a column of `width` inserted at `index`, keeping the total width:
 * the space is trimmed from the widest columns first (see shaveFromTop).
 */
export function takeSpace(widths: number[], index: number, width: number): number[] {
  const total = widths.reduce((a, b) => a + b, 0);
  const out = shaveFromTop(widths, width, () => true).map((w) => Math.round(w));
  out.splice(index, 0, Math.round(width));
  // Rounding: give the remainder to the widest column so the table width is exact.
  const diff = total - out.reduce((a, b) => a + b, 0);
  const widest = out.indexOf(Math.max(...out));
  out[widest] = (out[widest] as number) + diff;
  return out;
}

/**
 * Grid with the columns [start, start+span) resized to `target` in total (split in proportion
 * to their current widths). Growing trims the widest other columns first (shaveFromTop);
 * shrinking gives the freed space to the widest other column. The table width does not change.
 */
export function resizeSpan(widths: number[], start: number, span: number, target: number): number[] {
  const inSpan = (i: number) => i >= start && i < start + span;
  const current = widths.slice(start, start + span).reduce((a, b) => a + b, 0);
  const delta = target - current;
  let out = widths.slice();
  const others = widths.map((w, i) => (inSpan(i) ? 0 : w));
  if (delta > 0) out = shaveFromTop(widths, delta, (i) => !inSpan(i)).map((w) => Math.round(w));
  for (let i = start; i < start + span; i++) {
    out[i] = Math.round(current ? ((widths[i] as number) / current) * target : target / span);
  }
  if (delta < 0) {
    const widest = others.indexOf(Math.max(...others));
    out[widest] = (out[widest] as number) - delta;
  }
  // Rounding: keep the table width exact on the widest other column.
  const diff = widths.reduce((a, b) => a + b, 0) - out.reduce((a, b) => a + b, 0);
  const widestOther = others.indexOf(Math.max(...others));
  out[widestOther] = (out[widestOther] as number) + diff;
  return out;
}

/** Child order of w:rPr and w:pPr in the WordprocessingML schema; Word rejects out-of-order properties. */
const RPR_ORDER = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
];
const PPR_ORDER = [
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
];

/** Sets (replaces) the property `name` in a w:rPr / w:pPr, at its schema position. */
function setOrdered(props: Element, name: string, attrs: Record<string, string>, order: string[]): void {
  const doc = props.ownerDocument as Document;
  const old = child(props, name);
  if (old) props.removeChild(old);
  const el = doc.createElementNS(NS.w, `w:${name}`);
  for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(NS.w, k, v);
  const rank = order.indexOf(name);
  const before = children(props).find((c) => order.indexOf(localName(c)) > rank) ?? null;
  props.insertBefore(el, before);
}

/** Every cell's w:tcW recomputed from the grid columns it spans. */
function syncCellWidths(tbl: Element): void {
  const grid = children(child(tbl, 'tblGrid') as Element, 'gridCol').map((g) =>
    Number(g.getAttributeNS(NS.w, 'w') || 0),
  );
  for (const tr of tableRows(tbl)) {
    let col = 0;
    for (const { tc } of cellUnits(tr)) {
      const span = cellSpan(tc);
      const pr = child(tc, 'tcPr');
      const tcW = pr ? child(pr, 'tcW') : null;
      if (tcW && tcW.getAttributeNS(NS.w, 'type') !== 'pct') {
        tcW.setAttributeNS(NS.w, 'w:w', String(grid.slice(col, col + span).reduce((a, b) => a + b, 0)));
        tcW.setAttributeNS(NS.w, 'w:type', 'dxa');
      }
      col += span;
    }
  }
}

/** The direct cell "units" of a row: a w:tc, or a content control wrapping one. */
function cellUnits(tr: Element): Array<{ unit: Element; tc: Element }> {
  const out: Array<{ unit: Element; tc: Element }> = [];
  const walk = (container: Element, unit: Element | null) => {
    for (const c of children(container)) {
      if (localName(c) === 'tc') out.push({ unit: unit ?? c, tc: c });
      else if (localName(c) === 'sdt') walk(child(c, 'sdtContent') as Element, unit ?? c);
    }
  };
  walk(tr, null);
  return out;
}

function cellSpan(tc: Element): number {
  const pr = child(tc, 'tcPr');
  const gs = pr ? child(pr, 'gridSpan') : null;
  return Number(gs?.getAttributeNS(NS.w, 'val') || 1);
}

function setSpan(tc: Element, span: number): void {
  let pr = child(tc, 'tcPr');
  if (!pr) {
    pr = (tc.ownerDocument as Document).createElementNS(NS.w, 'w:tcPr');
    tc.insertBefore(pr, tc.firstChild);
  }
  let gs = child(pr, 'gridSpan');
  if (!gs) {
    gs = (tc.ownerDocument as Document).createElementNS(NS.w, 'w:gridSpan');
    const w = child(pr, 'tcW');
    pr.insertBefore(gs, w ? w.nextSibling : pr.firstChild);
  }
  gs.setAttributeNS(NS.w, 'w:val', String(span));
}

function rowGridWidth(tr: Element): number {
  return cellUnits(tr).reduce((n, c) => n + cellSpan(c.tc), 0);
}

function cellGridStart(tr: Element, tc: Element): number {
  let col = 0;
  for (const c of cellUnits(tr)) {
    if (c.tc === tc) return col;
    col += cellSpan(c.tc);
  }
  return col;
}

function cellCovering(tr: Element, gridIndex: number): Element | null {
  let col = 0;
  for (const c of cellUnits(tr)) {
    const span = cellSpan(c.tc);
    if (gridIndex >= col && gridIndex < col + span) return c.tc;
    col += span;
  }
  return null;
}

function cellAtGrid(tr: Element, gridIndex: number): Element | null {
  return cellCovering(tr, gridIndex);
}

/** The unit to insert before so the new cell lands at `gridIndex` (null → append). */
function cellUnitAtGrid(tr: Element, gridIndex: number): Element | null {
  let col = 0;
  for (const c of cellUnits(tr)) {
    if (col >= gridIndex) return c.unit;
    col += cellSpan(c.tc);
  }
  return null;
}

/** New cells look like their neighbour: same shading, borders, paragraph and run formatting. */
function copyCellLook(node: Element, neighbour: Element, width: number, align?: string): void {
  const tc = localName(node) === 'tc' ? node : (node.getElementsByTagNameNS(NS.w, 'tc')[0] as Element | undefined);
  if (!tc) return;
  const doc = tc.ownerDocument as Document;
  const nPr = child(neighbour, 'tcPr');
  if (nPr) {
    const copy = nPr.cloneNode(true) as Element;
    for (const x of children(copy)) if (['gridSpan', 'vMerge', 'hMerge'].includes(localName(x))) copy.removeChild(x);
    let w = child(copy, 'tcW');
    if (!w) {
      w = doc.createElementNS(NS.w, 'w:tcW');
      copy.insertBefore(w, copy.firstChild);
    }
    w.setAttributeNS(NS.w, 'w:w', String(width));
    w.setAttributeNS(NS.w, 'w:type', 'dxa');
    const old = child(tc, 'tcPr');
    if (old) tc.replaceChild(copy, old);
    else tc.insertBefore(copy, tc.firstChild);
  }
  const nP = children(neighbour, 'p')[0];
  const nPPr = nP ? child(nP, 'pPr') : null;
  const nRPr = nP ? Array.from(nP.getElementsByTagNameNS(NS.w, 'rPr'))[0] : undefined;
  for (const p of Array.from(tc.getElementsByTagNameNS(NS.w, 'p'))) {
    if (nPPr && !child(p, 'pPr')) p.insertBefore(nPPr.cloneNode(true), p.firstChild);
    if (align) {
      let pPr = child(p, 'pPr');
      if (!pPr) {
        pPr = doc.createElementNS(NS.w, 'w:pPr');
        p.insertBefore(pPr, p.firstChild);
      }
      const old = child(pPr, 'jc');
      if (old) pPr.removeChild(old);
      const jc = doc.createElementNS(NS.w, 'w:jc');
      jc.setAttributeNS(NS.w, 'w:val', align);
      pPr.appendChild(jc);
    }
    for (const r of Array.from(p.getElementsByTagNameNS(NS.w, 'r'))) {
      if (nRPr && !child(r, 'rPr')) r.insertBefore(nRPr.cloneNode(true), r.firstChild);
    }
  }
}

function isBefore(a: Node, b: Node): boolean {
  // compareDocumentPosition is not in xmldom: compare by walking the document in order.
  let found: 'a' | 'b' | null = null;
  const walk = (n: Node | null): void => {
    for (let c = n?.firstChild ?? null; c && !found; c = c.nextSibling) {
      if (c === a) found = 'a';
      else if (c === b) found = 'b';
      else walk(c);
    }
  };
  walk(a.ownerDocument);
  return found === 'a';
}

function setFirstText(sdt: Element, text: string): void {
  const ts = Array.from((child(sdt, 'sdtContent') as Element).getElementsByTagNameNS(NS.w, 't'));
  const first = ts[0];
  if (!first) return;
  while (first.firstChild) first.removeChild(first.firstChild);
  first.appendChild((first.ownerDocument as Document).createTextNode(text));
  for (const t of ts.slice(1)) t.parentNode?.removeChild(t);
}

/** Docentric elements directly inside a node (not nested in another element). */
function topLevelElementIds(node: Element): string[] {
  const own = localName(node) === 'sdt' ? sdtTag(node) : null;
  if (own) return [own.elementId];
  const out: string[] = [];
  const walk = (n: Element) => {
    for (const c of children(n)) {
      const t = localName(c) === 'sdt' ? sdtTag(c) : null;
      if (t) out.push(t.elementId);
      else walk(c);
    }
  };
  walk(node);
  return out;
}

function describeAnchor(a: { element: string } | { text: string }): string {
  return 'element' in a ? `#${a.element}` : `"${a.text}"`;
}

function describeNode(n: Element): string {
  const name = localName(n);
  return name === 'p' ? 'the paragraph' : name === 'tr' ? 'the table row' : name === 'sdt' ? 'the element' : name;
}

export { TemplateError };
