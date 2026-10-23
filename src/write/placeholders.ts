/**
 * The `{{…}}` placeholder syntax shared by the create command (in the layout spec)
 * and edit (typed by a person directly in Word):
 *
 *   {{PurchOrderHeader/@PurchId}}            Field, default data source of the scope
 *   {{GeneralData: CurrentCompany/@Name}}    Field, explicit data source
 *   {{@Qty * @PurchPrice | n2}}              Field with a .NET format string
 *   {{label:Total}}                          Label @DOC_LABEL:Total
 *   {{label:@SYS12345}}                      Label with a D365 label id
 *   {{sum: @Qty * @PurchPrice | n2}}         SummaryField over the enclosing table's rows
 *
 * The data-source prefix is recognised only for names the template declares,
 * so an XPath that happens to contain a colon is never misread.
 */

export type Segment =
  | { kind: 'text'; text: string }
  | { kind: 'field'; source: string | null; expr: string; format: string | null }
  | { kind: 'label'; code: string }
  | { kind: 'sum'; expr: string; format: string | null };

import { TEMPLATE_LABEL_PREFIX } from '../model/projectData.js';

const PLACEHOLDER = /\{\{([\s\S]+?)\}\}/g;
/** A trailing `| fmt` is a format only when it looks like one (n2, d, dd/MM/yyyy, #,##0.00 …). */
const FORMAT_SUFFIX = /\s*\|\s*([A-Za-z]{1,2}\d{0,2}|[#0.,]+|[dMyHhms/.\-: ]{2,})\s*$/;

export function parsePlaceholders(text: string, dataSources: string[]): Segment[] {
  const out: Segment[] = [];
  let last = 0;
  for (const m of text.matchAll(PLACEHOLDER)) {
    const start = m.index ?? 0;
    if (start > last) out.push({ kind: 'text', text: text.slice(last, start) });
    out.push(parseOne((m[1] as string).trim(), dataSources));
    last = start + m[0].length;
  }
  if (last < text.length) out.push({ kind: 'text', text: text.slice(last) });
  return out;
}

export function hasPlaceholders(text: string): boolean {
  return /\{\{[\s\S]+?\}\}/.test(text);
}

function parseOne(body: string, dataSources: string[]): Segment {
  const label = /^label\s*:\s*(.+)$/i.exec(body);
  if (label) {
    const code = (label[1] as string).trim();
    return { kind: 'label', code: code.startsWith('@') ? code : `${TEMPLATE_LABEL_PREFIX}${code}` };
  }
  const sum = /^sum\s*:\s*([\s\S]+)$/i.exec(body);
  if (sum) {
    const { expr, format } = splitFormat((sum[1] as string).trim());
    return { kind: 'sum', expr, format };
  }
  let source: string | null = null;
  let rest = body;
  const prefixed = /^([A-Za-z_][\w]*)\s*:\s*([\s\S]+)$/.exec(body);
  if (prefixed && dataSources.includes(prefixed[1] as string)) {
    source = prefixed[1] as string;
    rest = (prefixed[2] as string).trim();
  }
  const { expr, format } = splitFormat(rest);
  return { kind: 'field', source, expr, format };
}

function splitFormat(s: string): { expr: string; format: string | null } {
  const m = FORMAT_SUFFIX.exec(s);
  if (!m) return { expr: s, format: null };
  return { expr: s.slice(0, m.index).trim(), format: (m[1] as string).trim() };
}
