import { z } from 'zod';

/**
 * Layout specification: what an agent writes to create a Docentric template.
 * It describes structure (paragraphs, tables bound to collections, conditional
 * blocks) and leaves pixel-level layout to Word. Text accepts `{{…}}`
 * placeholders (see write/placeholders.ts). Full reference: docs/LAYOUT_SPEC.md.
 */

const hex = z.string().regex(/^#?[0-9A-Fa-f]{6}$/, 'RRGGBB hex colour');

export const TextStyle = z
  .object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    caps: z.boolean().optional(),
    size: z.number().min(4).max(96).optional(),
    color: hex.optional(),
    font: z.string().optional(),
    align: z.enum(['left', 'center', 'right', 'justify']).optional(),
    spaceAfter: z.number().min(0).max(100).optional(),
  })
  .strict();
export type TextStyle = z.infer<typeof TextStyle>;

const Source = z.string().describe("Data source name (MainData, GeneralData, Parameters) or '.' for the current row");

export type Block =
  | { text: string; style?: TextStyle }
  | { heading: string; level?: 1 | 2 | 3; style?: TextStyle }
  | { grid: Cell[][]; widths?: number[]; borders?: boolean; shade?: string; style?: TextStyle }
  | { columns: Block[][]; widths?: number[] }
  | { table: DataTable }
  | { if: string; source?: string; then: Block[] }
  | { list: string; source?: string; body: Block[] }
  | { var: string; value: string; source?: string }
  | { image: string; source?: string; width?: number; height?: number; align?: 'left' | 'center' | 'right' }
  | { pageBreak: true }
  | { spacer: true | number };

export type Cell = string | { blocks: Block[] } | { text: string; style?: TextStyle; shade?: string; span?: number };

export interface DataTable {
  rows: string;
  source?: string;
  columns: Array<{
    header?: string;
    value: string;
    width?: number;
    align?: 'left' | 'center' | 'right';
    format?: string;
  }>;
  headerStyle?: TextStyle;
  headerShade?: string;
  totals?: Cell[][];
  borders?: boolean;
  /** Show the table only when this condition holds (wraps it in an If). */
  if?: string;
}

const CellSchema: z.ZodType<Cell> = z.lazy(() =>
  z.union([
    z.string(),
    z.object({ blocks: z.array(BlockSchema) }).strict(),
    z
      .object({
        text: z.string(),
        style: TextStyle.optional(),
        shade: hex.optional(),
        span: z.number().int().min(1).optional(),
      })
      .strict(),
  ]),
);

const DataTableSchema = z
  .object({
    rows: z.string().min(1).describe('Collection path, e.g. PurchOrderLines'),
    source: Source.optional(),
    columns: z
      .array(
        z
          .object({
            header: z.string().optional(),
            value: z.string(),
            width: z.number().positive().optional(),
            align: z.enum(['left', 'center', 'right']).optional(),
            format: z.string().optional(),
          })
          .strict(),
      )
      .min(1),
    headerStyle: TextStyle.optional(),
    headerShade: hex.optional(),
    totals: z.array(z.array(CellSchema)).optional(),
    borders: z.boolean().optional(),
    if: z.string().optional(),
  })
  .strict();

/**
 * Blocks are told apart by their key ({text}, {table}, {if} …). Dispatching on
 * the key, rather than a plain union, gives the agent one precise error instead
 * of eleven "invalid input" alternatives.
 */
const BLOCK_SCHEMAS: Record<string, z.ZodType> = {
  text: z.object({ text: z.string(), style: TextStyle.optional() }).strict(),
  heading: z
    .object({
      heading: z.string(),
      level: z.union([z.literal(1), z.literal(2), z.literal(3)]).optional(),
      style: TextStyle.optional(),
    })
    .strict(),
  grid: z
    .object({
      grid: z.array(z.array(CellSchema)).min(1),
      widths: z.array(z.number().positive()).optional(),
      borders: z.boolean().optional(),
      shade: hex.optional(),
      style: TextStyle.optional(),
    })
    .strict(),
  columns: z
    .object({
      columns: z.array(z.array(z.lazy(() => BlockSchema))).min(1),
      widths: z.array(z.number().positive()).optional(),
    })
    .strict(),
  table: z.object({ table: DataTableSchema }).strict(),
  // biome-ignore lint/suspicious/noThenProperty: `then` is the spec's own key ({if, then}), never awaited.
  if: z.object({ if: z.string().min(1), source: Source.optional(), then: z.array(z.lazy(() => BlockSchema)) }).strict(),
  list: z
    .object({ list: z.string().min(1), source: Source.optional(), body: z.array(z.lazy(() => BlockSchema)) })
    .strict(),
  var: z
    .object({ var: z.string().regex(/^[A-Za-z_][\w-]*$/), value: z.string().min(1), source: Source.optional() })
    .strict(),
  image: z
    .object({
      image: z.string().min(1),
      source: Source.optional(),
      width: z.number().positive().optional(),
      height: z.number().positive().optional(),
      align: z.enum(['left', 'center', 'right']).optional(),
    })
    .strict(),
  pageBreak: z.object({ pageBreak: z.literal(true) }).strict(),
  spacer: z.object({ spacer: z.union([z.literal(true), z.number().positive()]) }).strict(),
};

export const BlockSchema: z.ZodType<Block> = z.unknown().superRefine((v, ctx) => {
  if (!v || typeof v !== 'object' || Array.isArray(v)) {
    ctx.addIssue({ code: 'custom', message: 'a block is an object such as {"text": …}, {"table": …} or {"if": …}' });
    return;
  }
  const key = Object.keys(BLOCK_SCHEMAS).find((k) => k in v);
  if (!key) {
    ctx.addIssue({
      code: 'custom',
      message: `unknown block {${Object.keys(v).join(', ')}}; expected one of: ${Object.keys(BLOCK_SCHEMAS).join(', ')}`,
    });
    return;
  }
  const r = (BLOCK_SCHEMAS[key] as z.ZodType).safeParse(v);
  if (!r.success) {
    for (const issue of r.error.issues)
      ctx.addIssue({ code: 'custom', message: issue.message, path: issue.path as (string | number)[] });
  }
}) as unknown as z.ZodType<Block>;

const AttrType = z.enum(['string', 'decimal', 'int', 'date', 'dateTime', 'boolean', 'base64Binary']);
export type AttrType = z.infer<typeof AttrType>;

export interface ElementShape {
  repeated?: boolean;
  attributes?: Record<string, AttrType>;
  elements?: Record<string, ElementShape>;
}

const ElementShapeSchema: z.ZodType<ElementShape> = z.lazy(() =>
  z
    .object({
      repeated: z.boolean().optional(),
      attributes: z.record(z.string(), AttrType).optional(),
      elements: z.record(z.string(), ElementShapeSchema).optional(),
    })
    .strict(),
);

export const DataSourceShape = z
  .object({
    root: z.string().regex(/^[A-Za-z_][\w.-]*$/),
    section: z.enum(['MainData', 'GeneralData', 'Parameters']).optional(),
    attributes: z.record(z.string(), AttrType).optional(),
    elements: z.record(z.string(), ElementShapeSchema).optional(),
    /** Sample XML to embed instead of generated values. */
    sample: z.string().optional(),
  })
  .strict();
export type DataSourceShape = z.infer<typeof DataSourceShape>;

export const LayoutSpec = z
  .object({
    title: z.string().optional(),
    report: z
      .object({
        id: z.string().optional(),
        dspClass: z.string().optional(),
        group: z.string().optional(),
      })
      .strict()
      .optional(),
    languages: z.array(z.string().min(2)).min(1).optional(),
    /** Template labels by code (without @DOC_LABEL:), text per language. */
    labels: z.record(z.string(), z.record(z.string(), z.string())).optional(),
    dataSources: z.union([
      z.object({ from: z.string().min(1) }).strict(),
      z.object({ schema: z.record(z.string(), DataSourceShape) }).strict(),
    ]),
    defaultSource: z.string().optional(),
    page: z
      .object({
        size: z.enum(['A4', 'Letter', 'A5']).optional(),
        orientation: z.enum(['portrait', 'landscape']).optional(),
        /** Millimetres: one number for all sides, or [top, right, bottom, left]. */
        margins: z.union([z.number().min(0), z.array(z.number().min(0)).length(4)]).optional(),
      })
      .strict()
      .optional(),
    font: z
      .object({ name: z.string().optional(), size: z.number().min(4).max(48).optional() })
      .strict()
      .optional(),
    /** Existing .docx whose styles, theme and fonts are reused (corporate look). */
    baseDocument: z.string().optional(),
    header: z.array(BlockSchema).optional(),
    footer: z.array(BlockSchema).optional(),
    body: z.array(BlockSchema).min(1),
  })
  .strict();
export type LayoutSpec = z.infer<typeof LayoutSpec>;

export function parseLayoutSpec(input: unknown): LayoutSpec {
  const r = LayoutSpec.safeParse(input);
  if (!r.success) {
    const issues = r.error.issues
      .slice(0, 12)
      .map((i) => `- ${i.path.join('.') || '(root)'}: ${i.message}`)
      .join('\n');
    throw new SpecError(`Invalid layout spec:\n${issues}`);
  }
  return r.data;
}

export class SpecError extends Error {}
