import {
  cell,
  inlinePicture,
  mmToTwips,
  type ParagraphStyle,
  paragraph,
  type RunStyle,
  row,
  run,
  sdt,
  table,
} from '../write/ooxml.js';
import { parsePlaceholders, type Segment } from '../write/placeholders.js';
import { IdAllocator, type NewElement } from '../write/projectDataXml.js';
import { type Block, type Cell, type DataTable, SpecError, type TextStyle } from './spec.js';

/**
 * Turns layout blocks into WordprocessingML plus the Docentric element registry.
 * Every element is registered under the element that contains it (ChildElements),
 * exactly as the Designer does, so the engine resolves Source='.' correctly.
 */

export const IMAGE_REL_TOKEN = '__DOCENTRIC_IMAGE__';

export interface Scope {
  /** Source for bindings without an explicit one: a data source name, or '.' inside a List. */
  source: string;
  parentId: string | null;
  /** Set inside a data table's totals: the collection {{sum:}} aggregates. */
  rows?: { source: string; path: string };
  /** Available width in twips, for tables. */
  width: number;
}

export class LayoutCompiler {
  readonly elements: NewElement[] = [];
  private readonly byId = new Map<string, NewElement>();
  imageCount = 0;

  constructor(
    private readonly projectId: string,
    private readonly ids: IdAllocator,
    private readonly dataSources: string[],
  ) {}

  blocks(list: Block[], scope: Scope): string {
    return list.map((b) => this.block(b, scope)).join('');
  }

  block(b: Block, scope: Scope): string {
    if ('text' in b) return this.textParagraph(b.text, scope, b.style);
    if ('heading' in b) {
      const level = b.level ?? 1;
      return paragraph(this.inline(b.heading, scope, runStyle(b.style)), {
        ...paraStyle(b.style),
        styleId: `Heading${level}`,
      });
    }
    if ('grid' in b) return this.grid(b.grid, b.widths, b.borders ?? false, b.shade, b.style, scope);
    if ('columns' in b) {
      return this.grid([b.columns.map((blocks) => ({ blocks }))], b.widths, false, undefined, undefined, scope);
    }
    if ('table' in b) return this.dataTable(b.table, scope);
    if ('if' in b) {
      const el = this.register({ type: 'If', source: b.source ?? scope.source, path: b.if, children: [] }, scope);
      const inner = this.blocks(b.then, { ...scope, parentId: el.id }) || paragraph('');
      return sdt(this.sdtOptions(el), inner);
    }
    if ('list' in b) {
      const el = this.register(
        { type: 'List', source: b.source ?? scope.source, path: b.list, attributes: { RenderMode: '1' }, children: [] },
        scope,
      );
      const inner = this.blocks(b.body, { ...scope, source: '.', parentId: el.id, rows: undefined }) || paragraph('');
      return sdt(this.sdtOptions(el), inner);
    }
    if ('var' in b) {
      const el = this.register(
        { type: 'Var', name: b.var, source: b.source ?? scope.source, path: b.value, children: [] },
        scope,
      );
      return paragraph(sdt(this.sdtOptions(el), run(`${b.var} = ${b.value}`, { color: '808080', size: 8 })));
    }
    if ('image' in b) {
      const el = this.register(
        { type: 'Image', source: b.source ?? scope.source, path: b.image, attributes: { SizeMode: '4' }, children: [] },
        scope,
      );
      this.imageCount++;
      const pic = inlinePicture(
        IMAGE_REL_TOKEN,
        b.width ?? 40,
        b.height ?? 20,
        100 + this.imageCount,
        `Image ${el.id}`,
      );
      return paragraph(sdt(this.sdtOptions(el), pic), { align: b.align });
    }
    if ('pageBreak' in b) return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    if ('spacer' in b) return paragraph('', { spaceAfter: typeof b.spacer === 'number' ? b.spacer : undefined });
    throw new SpecError(`Unknown block: ${JSON.stringify(b).slice(0, 80)}`);
  }

  /** A paragraph whose text may hold placeholders. */
  textParagraph(text: string, scope: Scope, style?: TextStyle): string {
    return paragraph(this.inline(text, scope, runStyle(style)), paraStyle(style), runStyle(style));
  }

  /** Like inline(), but every run carries a raw <w:rPr> copied from the text it replaces. */
  inlineWithRawRunProps(text: string, scope: Scope, rawRPr: string): string {
    const xml = this.inline(text, scope);
    return rawRPr ? xml.replace(/<w:r>/g, `<w:r>${rawRPr}`) : xml;
  }

  inline(text: string, scope: Scope, style?: RunStyle): string {
    return parsePlaceholders(text, this.dataSources)
      .map((seg) => (seg.kind === 'text' ? run(seg.text, style) : this.elementRun(seg, scope, style)))
      .join('');
  }

  private elementRun(seg: Exclude<Segment, { kind: 'text' }>, scope: Scope, style?: RunStyle, format?: string): string {
    const el = this.elementFor(seg, scope, format);
    return sdt(this.sdtOptions(el, style), run(displayText(seg), style));
  }

  private elementFor(seg: Exclude<Segment, { kind: 'text' }>, scope: Scope, format?: string): NewElement {
    if (seg.kind === 'label') return this.register({ type: 'Label', labelCode: seg.code, children: [] }, scope);
    if (seg.kind === 'sum') {
      if (!scope.rows) throw new SpecError(`{{sum: ${seg.expr}}} is only allowed in a table's "totals" rows.`);
      const fmt = seg.format ?? format;
      return this.register(
        {
          type: 'SummaryField',
          source: scope.rows.source,
          path: scope.rows.path,
          aggregatePath: seg.expr,
          attributes: { AggregateFunction: '1', IgnoreNullValues: '0', ...(fmt ? { FormatString: fmt } : {}) },
          children: [],
        },
        scope,
      );
    }
    const fmt = seg.format ?? format;
    return this.register(
      {
        type: 'Field',
        source: seg.source ?? scope.source,
        path: seg.expr,
        attributes: fmt ? { FormatString: fmt } : {},
        children: [],
      },
      scope,
    );
  }

  // --- Tables --------------------------------------------------------------------

  private grid(
    rows: Cell[][],
    weights: number[] | undefined,
    borders: boolean,
    shade: string | undefined,
    style: TextStyle | undefined,
    scope: Scope,
  ): string {
    const cols = Math.max(
      ...rows.map((r) => r.reduce((n, c) => n + (typeof c === 'object' && 'span' in c && c.span ? c.span : 1), 0)),
    );
    const widths = distribute(scope.width, cols, weights);
    const xmlRows = rows
      .map((r) => {
        let col = 0;
        const cells = r
          .map((c) => {
            const span = typeof c === 'object' && 'span' in c && c.span ? c.span : 1;
            const w = widths.slice(col, col + span).reduce((a, b) => a + b, 0);
            col += span;
            return this.cellXml(c, scope, { width: w, span, borders, shade, style });
          })
          .join('');
        return row(cells);
      })
      .join('');
    return `${table(xmlRows, widths, { borders, layoutFixed: true })}${paragraph('')}`;
  }

  private dataTable(t: DataTable, outer: Scope): string {
    let scope = outer;
    let wrapIf: NewElement | null = null;
    if (t.if) {
      wrapIf = this.register({ type: 'If', source: outer.source, path: t.if, children: [] }, outer);
      scope = { ...outer, parentId: wrapIf.id };
    }
    const listSource = t.source ?? scope.source;
    const widths = distribute(
      scope.width,
      t.columns.length,
      t.columns.map((c) => c.width ?? 1),
    );
    const borders = t.borders ?? true;
    const headerStyle: TextStyle = { bold: true, ...t.headerStyle };
    const headerShade = t.headerShade ?? 'F2F2F2';

    let xml = '';
    if (t.columns.some((c) => c.header)) {
      const cells = t.columns
        .map((c, i) =>
          this.cellXml(c.header ?? '', scope, {
            width: widths[i] as number,
            borders,
            shade: headerShade,
            style: { ...headerStyle, align: c.align },
          }),
        )
        .join('');
      xml += row(cells, { header: true, cantSplit: true });
    }

    const list = this.register(
      { type: 'List', source: listSource, path: t.rows, attributes: { RenderMode: '2' }, children: [] },
      scope,
    );
    const rowScope: Scope = { source: '.', parentId: list.id, width: scope.width };
    const dataCells = t.columns
      .map((c, i) =>
        this.cellXml(c.value, rowScope, { width: widths[i] as number, borders, style: { align: c.align } }, c.format),
      )
      .join('');
    xml += sdt(this.sdtOptions(list), row(dataCells, { cantSplit: true }));

    const totalsScope: Scope = { ...scope, rows: { source: listSource, path: t.rows } };
    for (const totalRow of t.totals ?? []) {
      let col = 0;
      const cells = totalRow
        .map((c) => {
          const span = typeof c === 'object' && 'span' in c && c.span ? c.span : 1;
          const w = widths.slice(col, col + span).reduce((a, b) => a + b, 0);
          const align = t.columns[col]?.align;
          col += span;
          return this.cellXml(c, totalsScope, { width: w, span, borders, style: { bold: true, align } });
        })
        .join('');
      xml += row(cells, { cantSplit: true });
    }

    const tbl = table(xml, widths, { borders, layoutFixed: true });
    const out = wrapIf ? sdt(this.sdtOptions(wrapIf), `${tbl}${paragraph('')}`) : `${tbl}${paragraph('')}`;
    return out;
  }

  /**
   * One cell. A cell holding exactly one placeholder becomes a cell-level content
   * control, as the Designer makes it; anything else is a normal cell.
   */
  cellXml(
    c: Cell,
    scope: Scope,
    o: { width: number; span?: number; borders: boolean; shade?: string; style?: TextStyle },
    format?: string,
  ): string {
    const cellOpts = { width: o.width, span: o.span, shade: o.shade, borders: o.borders ? undefined : false };
    if (typeof c === 'object' && 'blocks' in c) {
      return cell(this.blocks(c.blocks, { ...scope, width: o.width - 170 }), cellOpts);
    }
    const text = typeof c === 'string' ? c : c.text;
    const style: TextStyle = { ...o.style, ...(typeof c === 'object' ? c.style : {}) };
    if (typeof c === 'object' && c.shade) cellOpts.shade = c.shade;
    const segs = parsePlaceholders(text, this.dataSources);
    const only = segs.length === 1 ? segs[0] : undefined;
    if (only && only.kind !== 'text') {
      const el = this.elementFor(only, scope, format);
      const inner = paragraph(run(displayText(only), runStyle(style)), paraStyle(style), runStyle(style));
      return sdt(this.sdtOptions(el, runStyle(style)), cell(inner, cellOpts));
    }
    const inner = paragraph(
      segs
        .map((s) =>
          s.kind === 'text' ? run(s.text, runStyle(style)) : this.elementRun(s, scope, runStyle(style), format),
        )
        .join(''),
      paraStyle(style),
      runStyle(style),
    );
    return cell(inner, cellOpts);
  }

  // --- Registry ------------------------------------------------------------------

  private register(e: Omit<NewElement, 'id'>, scope: Scope): NewElement {
    const el: NewElement = { id: this.ids.elementId(), ...e };
    this.elements.push(el);
    this.byId.set(el.id, el);
    if (scope.parentId) this.byId.get(scope.parentId)?.children.push(el.id);
    return el;
  }

  private sdtOptions(el: NewElement, runStyle?: RunStyle) {
    return { alias: el.type, tag: `${this.projectId}-${el.id}`, id: IdAllocator.sdtId(), runStyle };
  }
}

function displayText(seg: Exclude<Segment, { kind: 'text' }>): string {
  if (seg.kind === 'label') return seg.code;
  if (seg.kind === 'sum') return `Σ ${seg.expr}`;
  return seg.expr;
}

export function runStyle(s: TextStyle | undefined): RunStyle | undefined {
  if (!s) return undefined;
  const r: RunStyle = {
    bold: s.bold,
    italic: s.italic,
    underline: s.underline,
    caps: s.caps,
    size: s.size,
    color: s.color,
    font: s.font,
  };
  return Object.values(r).some((v) => v !== undefined) ? r : undefined;
}

export function paraStyle(s: TextStyle | undefined): ParagraphStyle | undefined {
  if (!s) return undefined;
  if (s.align === undefined && s.spaceAfter === undefined) return undefined;
  return { align: s.align, spaceAfter: s.spaceAfter };
}

/** Splits `total` twips across `count` columns by weight (equal when none given). */
export function distribute(total: number, count: number, weights?: number[]): number[] {
  const w = Array.from({ length: count }, (_, i) => weights?.[i] ?? 1);
  const sum = w.reduce((a, b) => a + b, 0);
  return w.map((x) => Math.floor((x / sum) * total));
}

export const PAGE_SIZES: Record<string, [number, number]> = {
  A4: [11906, 16838],
  Letter: [12240, 15840],
  A5: [8391, 11906],
};

export function pageGeometry(page: { size?: string; orientation?: string; margins?: number | number[] } | undefined) {
  const [w0, h0] = PAGE_SIZES[page?.size ?? 'A4'] ?? [11906, 16838];
  const landscape = page?.orientation === 'landscape';
  const [w, h] = landscape ? [h0, w0] : [w0, h0];
  const m = page?.margins ?? 18;
  const [top, right, bottom, left] = (Array.isArray(m) ? m : [m, m, m, m]).map(mmToTwips) as [
    number,
    number,
    number,
    number,
  ];
  return { width: w, height: h, landscape, top, right, bottom, left, contentWidth: w - left - right };
}
