import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { Element } from '@xmldom/xmldom';
import type { AttrType, Block, Cell, DataSourceShape, LayoutSpec, TextStyle } from '../create/spec.js';
import { assertAllowedPath, parseDocument, TemplateError } from '../docx/package.js';
import { child, children, descendants, localName } from '../xml/dom.js';

/**
 * SSRS (AxReport) → Docentric layout spec. A starting point, not a pixel copy:
 * absolute RDL positions become rows and columns, Tablix become data tables,
 * `=Fields!X.Value` becomes `{{@X}}`. Everything that cannot be translated
 * safely is listed in `notes` and shown in the layout as ⟦SSRS: …⟧ so the agent
 * (or a person) can finish it.
 *
 * The schema is derived from the report datasets and is provisional: the real
 * element names come from the Docentric DDSP generated in D365.
 */

export interface ConvertResult {
  spec: LayoutSpec;
  notes: string[];
  stats: { textboxes: number; tables: number; images: number; untranslated: number };
}

export interface ConvertOptions {
  design?: string;
}

const TYPE_MAP: Record<string, AttrType> = {
  'System.String': 'string',
  'System.Double': 'decimal',
  'System.Decimal': 'decimal',
  'System.Single': 'decimal',
  'System.Int16': 'int',
  'System.Int32': 'int',
  'System.Int64': 'int',
  'System.DateTime': 'date',
  'System.Boolean': 'boolean',
  'System.Byte[]': 'base64Binary',
};

interface Dataset {
  name: string;
  fields: Map<string, AttrType>;
}

export async function convertSsrs(axReportPath: string, opts: ConvertOptions = {}): Promise<ConvertResult> {
  const full = resolve(axReportPath);
  assertAllowedPath(full);
  const xml = await readFile(full, 'utf8').catch(() => {
    throw new TemplateError(`File not found: ${full}`);
  });
  return convertAxReport(xml, opts);
}

export function convertAxReport(axReportXml: string, opts: ConvertOptions = {}): ConvertResult {
  const ax = parseDocument(axReportXml, 'AxReport').documentElement as Element;
  if (localName(ax) !== 'AxReport')
    throw new TemplateError('Not an AxReport metadata file (root element is not AxReport).');
  const reportName = textOf(child(ax, 'Name')) || 'Report';

  const datasets = new Map<string, Dataset>();
  for (const ds of descendants(ax, 'AxReportDataSet')) {
    const name = textOf(child(ds, 'Name'));
    const fields = new Map<string, AttrType>();
    for (const f of descendants(ds, 'AxReportDataSetField')) {
      fields.set(textOf(child(f, 'Name')), TYPE_MAP[textOf(child(f, 'DataType'))] ?? 'string');
    }
    if (name) datasets.set(name, { name, fields });
  }

  const designs = descendants(ax, 'AxReportDesign');
  const design = (opts.design ? designs.find((d) => textOf(child(d, 'Name')) === opts.design) : designs[0]) ?? null;
  if (!design) throw new TemplateError(opts.design ? `No design named '${opts.design}'.` : 'The report has no design.');
  const designName = textOf(child(design, 'Name')) || 'Report';
  const rdlText = textOf(child(design, 'Text'));
  if (!rdlText.trim())
    throw new TemplateError(`Design '${designName}' has no RDL (auto-design reports are not supported).`);
  const rdl = parseDocument(rdlText.trim(), `${reportName}.${designName}.rdl`).documentElement as Element;

  const conv = new RdlConverter(datasets);
  const body = descendants(rdl, 'Body')[0];
  const header = descendants(rdl, 'PageHeader')[0];
  const footer = descendants(rdl, 'PageFooter')[0];

  // Layout first: it records which datasets are repeated, read once, or not used — the schema needs that.
  const headerBlocks = header ? conv.items(child(header, 'ReportItems'), null) : [];
  const bodyBlocks = body ? conv.items(child(body, 'ReportItems'), null) : [];
  const footerBlocks = footer ? conv.items(child(footer, 'ReportItems'), null) : [];
  const spec: LayoutSpec = {
    title: `${reportName} (${designName})`,
    report: { id: `${reportName}.${designName}` },
    languages: ['en-us'],
    dataSources: { schema: conv.schema(rdl) },
    header: headerBlocks,
    body: bodyBlocks as LayoutSpec['body'],
    footer: footerBlocks,
  };
  if (!spec.body.length) spec.body = [{ text: '' }];
  if (!spec.header?.length) delete spec.header;
  if (!spec.footer?.length) delete spec.footer;
  return { spec, notes: conv.notes, stats: conv.stats };
}

class RdlConverter {
  readonly notes: string[] = [];
  readonly stats = { textboxes: 0, tables: 0, images: 0, untranslated: 0 };
  /** Datasets used as table rows (repeated) vs read with First() (single). */
  private readonly repeated = new Set<string>();
  private readonly used = new Set<string>();
  private readonly parameters = new Map<string, AttrType>();
  private usesCompany = false;

  constructor(private readonly datasets: Map<string, Dataset>) {}

  schema(rdl: Element): Record<string, DataSourceShape> {
    for (const p of descendants(rdl, 'ReportParameter')) {
      const name = p.getAttribute('Name') ?? '';
      if (name && !name.startsWith('AX_')) {
        this.parameters.set(name, TYPE_MAP[`System.${textOf(child(p, 'DataType'))}`] ?? 'string');
      }
    }
    const elements: NonNullable<DataSourceShape['elements']> = {};
    for (const ds of this.datasets.values()) {
      elements[ds.name] = {
        repeated: this.repeated.has(ds.name) || !this.used.has(ds.name),
        attributes: Object.fromEntries(ds.fields),
      };
    }
    const out: Record<string, DataSourceShape> = { MainData: { root: 'ReportData', section: 'MainData', elements } };
    if (this.parameters.size) {
      out.Parameters = { root: 'Parameters', section: 'Parameters', attributes: Object.fromEntries(this.parameters) };
    }
    if (this.usesCompany) {
      out.GeneralData = {
        root: 'GeneralData',
        section: 'GeneralData',
        elements: { CurrentCompany: { attributes: { Name: 'string' } } },
      };
    }
    return out;
  }

  /** Report items laid out by position: items sharing a top edge form one row. */
  items(container: Element | null, rowDataset: string | null): Block[] {
    if (!container) return [];
    const placed = children(container)
      .map((el) => ({
        el,
        top: length(child(el, 'Top')),
        left: length(child(el, 'Left')),
        width: length(child(el, 'Width')),
      }))
      .sort((a, b) => a.top - b.top || a.left - b.left);
    const rows: (typeof placed)[] = [];
    for (const p of placed) {
      const last = rows[rows.length - 1];
      if (last && Math.abs((last[0]?.top ?? 0) - p.top) < 0.08) last.push(p);
      else rows.push([p]);
    }
    const blocks: Block[] = [];
    for (const r of rows) {
      if (r.length === 1) {
        blocks.push(...this.item((r[0] as (typeof placed)[number]).el, rowDataset));
        continue;
      }
      // Side by side items → a borderless grid, column widths from the RDL widths.
      const cells: Cell[] = r.map((p) => {
        const inner = this.item(p.el, rowDataset);
        const only = inner.length === 1 ? inner[0] : undefined;
        if (only && 'text' in only) return { text: only.text, style: only.style };
        return { blocks: inner.length ? inner : [{ text: '' }] };
      });
      blocks.push({ grid: [cells], widths: r.map((p) => Math.max(0.1, p.width || 1)), borders: false });
    }
    return blocks;
  }

  private item(el: Element, rowDataset: string | null): Block[] {
    switch (localName(el)) {
      case 'Textbox':
        this.stats.textboxes++;
        return [{ text: this.textboxText(el, rowDataset), style: textStyle(el) }];
      case 'Tablix':
        return this.tablix(el);
      case 'Rectangle':
        return this.items(child(el, 'ReportItems'), rowDataset);
      case 'Image':
        return this.image(el);
      case 'Line':
        return [];
      default:
        this.note(`${localName(el)} '${el.getAttribute('Name') ?? ''}' is not converted.`);
        return [{ text: `⟦SSRS ${localName(el)}: ${el.getAttribute('Name') ?? ''}⟧` }];
    }
  }

  private image(el: Element): Block[] {
    this.stats.images++;
    const value = textOf(child(el, 'Value'));
    const field = /^=\s*(?:First\()?Fields!(\w+)\.Value(?:\s*,\s*"(\w+)"\))?/.exec(value);
    if (textOf(child(el, 'Source')) === 'Database' && field) {
      const ds = field[2] ?? this.firstDatasetWith(field[1] as string);
      if (ds) this.used.add(ds);
      return [
        {
          image: ds ? `${ds}/@${field[1]}` : `@${field[1]}`,
          width: inches(length(child(el, 'Width'))),
          height: inches(length(child(el, 'Height'))),
        },
      ];
    }
    this.note(
      `Image '${el.getAttribute('Name') ?? ''}' (${textOf(child(el, 'Source')) || 'unknown source'}) needs a binding.`,
    );
    return [{ text: `⟦SSRS image: ${el.getAttribute('Name') ?? ''}⟧` }];
  }

  private tablix(t: Element): Block[] {
    this.stats.tables++;
    const dataset = textOf(child(t, 'DataSetName'));
    if (dataset) {
      this.repeated.add(dataset);
      this.used.add(dataset);
    }
    const body = child(t, 'TablixBody');
    const widths = descendants(child(body ?? t, 'TablixColumns') ?? t, 'TablixColumn').map(
      (c) => length(child(c, 'Width')) || 1,
    );
    const rows = descendants(child(body ?? t, 'TablixRows') ?? t, 'TablixRow');
    const kinds = rowKinds(t, rows);

    const cellsOf = (row: Element, scope: string | null) =>
      descendants(row, 'TablixCell').map((tc) => {
        const tb = descendants(tc, 'Textbox')[0];
        const span = Number(textOf(descendants(tc, 'ColSpan')[0] ?? null) || 1);
        const text = tb ? this.textboxText(tb, scope, dataset) : '';
        if (tb) this.stats.textboxes++;
        return { text, span, style: tb ? textStyle(tb) : undefined };
      });

    const header = rows.filter((_, i) => kinds[i] === 'header');
    const detail = rows.find((_, i) => kinds[i] === 'detail');
    const totals = rows.filter((_, i) => kinds[i] === 'total');
    const grouped = rows.filter((_, i) => kinds[i] === 'group').length;
    if (grouped)
      this.note(
        `Tablix '${t.getAttribute('Name')}': ${grouped} group row(s) dropped — Docentric groups need a Group element (add it in the Designer).`,
      );
    if (!detail) {
      this.note(`Tablix '${t.getAttribute('Name')}' has no detail row; converted as a static grid.`);
      return [
        {
          grid: rows.map((r) => cellsOf(r, null).map((c) => ({ text: c.text, span: c.span, style: c.style }))),
          widths,
          borders: true,
        },
      ];
    }

    const detailCells = cellsOf(detail, dataset);
    const headerCells = header.length ? cellsOf(header[header.length - 1] as Element, null) : [];
    const columns = detailCells.map((c, i) => ({
      header: headerCells[i]?.text || undefined,
      value: c.text || '',
      width: widths[i] ?? 1,
      align: c.style?.align === 'justify' ? undefined : c.style?.align,
    }));
    const table: Extract<Block, { table: unknown }>['table'] = {
      rows: dataset || 'Rows',
      columns: columns.map((c) => ({ ...c, value: c.value || ' ' })),
    };
    if (totals.length) {
      table.totals = totals.map((r) =>
        cellsOf(r, dataset).map((c) => ({ text: c.text.replace(/\{\{@/g, '{{sum: @'), span: c.span, style: c.style })),
      );
    }
    const blocks: Block[] = [];
    for (const extra of header.slice(0, -1)) {
      blocks.push({
        grid: [cellsOf(extra, null).map((c) => ({ text: c.text, span: c.span, style: c.style }))],
        borders: false,
      });
    }
    blocks.push({ table });
    return blocks;
  }

  /** The visible text of a textbox, with expressions turned into placeholders. */
  private textboxText(tb: Element, rowDataset: string | null, tableDataset?: string): string {
    const direct = child(tb, 'Value');
    const values = direct ? [direct] : descendants(tb, 'Value');
    return values.map((v) => this.expression(textOf(v), rowDataset, tableDataset)).join('');
  }

  expression(value: string, rowDataset: string | null, tableDataset?: string): string {
    const v = value.trim();
    if (!v.startsWith('=')) {
      if (/^@[A-Za-z]+:?\w+$/.test(v)) return `{{label:${v}}}`;
      const quoted = /^"((?:[^"]|"")*)"$/.exec(v);
      return quoted ? (quoted[1] as string).replace(/""/g, '"') : v;
    }
    const e = v.slice(1).trim();
    const translated = this.translate(e, rowDataset, tableDataset);
    if (translated !== null) return translated;
    this.stats.untranslated++;
    this.note(`Expression not translated: =${e}`);
    return `⟦SSRS: =${e}⟧`;
  }

  /** Translates a VB expression to placeholder text, or null when it cannot do so safely. */
  private translate(e: string, rowDataset: string | null, tableDataset?: string): string | null {
    const literal = /^"((?:[^"]|"")*)"$/.exec(e);
    if (literal) return (literal[1] as string).replace(/""/g, '"');
    const label = /^Labels!(@\w+:?\w*)$/.exec(e);
    if (label) return `{{label:${label[1]}}}`;
    const field = /^Fields!(\w+)\.Value$/.exec(e);
    if (field) return this.field(field[1] as string, rowDataset ?? tableDataset ?? null, !!rowDataset);
    const first = /^First\(Fields!(\w+)\.Value(?:\s*,\s*"(\w+)")?\)$/.exec(e);
    if (first) {
      const name = first[1] as string;
      return this.field(name, first[2] ?? tableDataset ?? this.firstDatasetWith(name), false);
    }
    const sum = /^Sum\(Fields!(\w+)\.Value(?:\s*,\s*"(\w+)")?\)$/.exec(e);
    if (sum) {
      if (rowDataset || tableDataset) return `{{@${sum[1]}}}`; // inside a table: becomes {{sum:}} in totals
      const ds = sum[2];
      return ds ? `{{sum(${ds}/@${sum[1]})}}` : null;
    }
    const param = /^Parameters!(\w+)\.Value$/.exec(e);
    if (param) {
      const name = param[1] as string;
      if (name === 'AX_CompanyName') {
        this.usesCompany = true;
        return '{{GeneralData: CurrentCompany/@Name}}';
      }
      if (name.startsWith('AX_')) return null;
      this.parameters.set(name, this.parameters.get(name) ?? 'string');
      return `{{Parameters: @${name}}}`;
    }
    const format = /^Format\((.+),\s*"([^"]+)"\)$/.exec(e);
    if (format) {
      const inner = this.translate((format[1] as string).trim(), rowDataset, tableDataset);
      return inner && /^\{\{[^{}]+\}\}$/.test(inner) ? inner.replace(/\}\}$/, ` | ${format[2]}}}`) : null;
    }
    // a & b & c → concatenation of translatable parts
    const parts = splitTopLevel(e, '&');
    if (parts.length > 1) {
      const out = parts.map((p) => this.translate(p.trim(), rowDataset, tableDataset));
      return out.every((x) => x !== null) ? out.join('') : null;
    }
    return null;
  }

  private field(name: string, dataset: string | null, inRow: boolean): string {
    if (inRow) return `{{@${name}}}`;
    if (!dataset) return `{{@${name}}}`;
    this.used.add(dataset);
    return `{{${dataset}/@${name}}}`;
  }

  private firstDatasetWith(field: string): string | null {
    for (const ds of this.datasets.values()) if (ds.fields.has(field)) return ds.name;
    return null;
  }

  private note(n: string): void {
    if (!this.notes.includes(n)) this.notes.push(n);
  }
}

/**
 * Classifies Tablix rows. Content decides first — a detail row reads plain
 * fields, a total row aggregates, a header row holds literals — because real
 * reports often have unconventional row hierarchies. The hierarchy (the member
 * whose Group has no GroupExpression) breaks ties between several detail-like rows.
 */
function rowKinds(t: Element, rows: Element[]): Array<'header' | 'detail' | 'group' | 'total'> {
  const values = rows.map((r) =>
    descendants(r, 'Value')
      .map((v) => v.textContent ?? '')
      .join(' '),
  );
  const AGG = /\b(?:Sum|First|Last|Count|Avg|Min|Max|CountDistinct)\([^)]*\)/g;
  const plainField = (v: string) => /(^|[^(\w])Fields!\w+\.Value/.test(v.replace(AGG, ''));
  const aggregate = (v: string) => /\b(?:Sum|Count|Avg|Min|Max|CountDistinct)\(/.test(v);
  const candidates = rows.map((_, i) => plainField(values[i] as string) && !aggregate(values[i] as string));

  // Hierarchy hint: leaf index of the Details member.
  const leaves: boolean[] = [];
  const walk = (members: Element | null) => {
    for (const m of members ? children(members, 'TablixMember') : []) {
      const group = child(m, 'Group');
      const sub = child(m, 'TablixMembers');
      if (sub) walk(sub);
      else leaves.push(!!group && /detail/i.test(group.getAttribute('Name') ?? ''));
    }
  };
  const hierarchy = child(t, 'TablixRowHierarchy');
  walk(hierarchy ? child(hierarchy, 'TablixMembers') : null);

  let detail = candidates.findIndex((c, i) => c && leaves[i]);
  if (detail < 0) detail = candidates.lastIndexOf(true);
  return rows.map((_, i) => {
    if (i === detail) return 'detail';
    if (detail < 0 || i < detail) return plainField(values[i] as string) ? 'group' : 'header';
    return 'total';
  });
}

function textStyle(el: Element): TextStyle | undefined {
  const style = child(el, 'Style');
  const run = descendants(el, 'TextRun')[0];
  const runStyle = run ? child(run, 'Style') : null;
  const para = descendants(el, 'Paragraph')[0];
  const paraStyle = para ? child(para, 'Style') : null;
  const get = (name: string) =>
    textOf(runStyle ? child(runStyle, name) : null) ||
    textOf(paraStyle ? child(paraStyle, name) : null) ||
    textOf(style ? child(style, name) : null);
  const s: TextStyle = {};
  if (/bold|semibold|heavy|extrabold/i.test(get('FontWeight'))) s.bold = true;
  if (/italic/i.test(get('FontStyle'))) s.italic = true;
  const size = Number.parseFloat(get('FontSize'));
  if (size) s.size = size;
  const align = get('TextAlign').toLowerCase();
  if (align === 'right' || align === 'center') s.align = align;
  const color = get('Color');
  if (/^#[0-9a-f]{6}$/i.test(color)) s.color = color.slice(1);
  return Object.keys(s).length ? s : undefined;
}

/** RDL sizes (in, cm, mm, pt) → inches. */
function length(el: Element | null): number {
  const s = textOf(el);
  const m = /^([\d.]+)\s*(in|cm|mm|pt)?$/.exec(s);
  if (!m) return 0;
  const n = Number(m[1]);
  switch (m[2]) {
    case 'cm':
      return n / 2.54;
    case 'mm':
      return n / 25.4;
    case 'pt':
      return n / 72;
    default:
      return n;
  }
}

const inches = (n: number) => (n ? Math.round(n * 25.4) : undefined);

function textOf(el: Element | null | undefined): string {
  return el?.textContent ?? '';
}

/** Splits on `sep` outside quotes and parentheses. */
function splitTopLevel(s: string, sep: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let quoted = false;
  let cur = '';
  for (const c of s) {
    if (c === '"') quoted = !quoted;
    else if (!quoted && c === '(') depth++;
    else if (!quoted && c === ')') depth--;
    if (!quoted && depth === 0 && c === sep) {
      out.push(cur);
      cur = '';
    } else cur += c;
  }
  out.push(cur);
  return out;
}
