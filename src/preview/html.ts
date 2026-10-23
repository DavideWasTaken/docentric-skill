import type { Element } from '@xmldom/xmldom';
import { NS, type WordPackage } from '../docx/package.js';
import { children, localName } from '../xml/dom.js';

/**
 * Minimal WordprocessingML → HTML renderer for previews. It keeps what matters to
 * judge a business document — text, bold/italic/size/colour, alignment, tables,
 * line breaks, headers and footers — and ignores the rest. Always available: no
 * Word, LibreOffice or browser needed.
 */

export async function renderHtml(pkg: WordPackage, title: string): Promise<string> {
  const docDom = await pkg.dom('word/document.xml');
  if (!docDom) return '';
  const body = children(docDom.documentElement as Element, 'body')[0];
  const rels = new Map((await pkg.relationships('word/document.xml')).map((r) => [r.id, r.target]));

  const sectPr = body ? children(body, 'sectPr')[0] : undefined;
  const headerHtml = sectPr ? await renderReferenced(pkg, sectPr, 'headerReference', rels) : '';
  const footerHtml = sectPr ? await renderReferenced(pkg, sectPr, 'footerReference', rels) : '';
  const page = pageBox(sectPr);

  return `<!doctype html>
<html data-page-width="${page.width}"><head><meta charset="utf-8"><title>${esc(title)}</title>
<style>
  body { background:#e9eaee; margin:0; padding:24px 0; font-family:Calibri,'Segoe UI',Arial,sans-serif; }
  .page { background:#fff; width:${page.width}px; min-height:${page.height}px; margin:0 auto;
          padding:${page.top}px ${page.right}px ${page.bottom}px ${page.left}px; box-sizing:border-box;
          box-shadow:0 2px 12px rgba(0,0,0,.15); font-size:11pt; color:#111; }
  .hdr { border-bottom:1px dashed #c8c8d0; margin-bottom:16px; padding-bottom:8px; }
  .ftr { border-top:1px dashed #c8c8d0; margin-top:24px; padding-top:8px; font-size:9pt; }
  p { margin:0 0 4px; min-height:1em; }
  table { border-collapse:collapse; width:100%; margin:4px 0 8px; }
  td { border:1px solid #d6d6de; padding:3px 5px; vertical-align:top; overflow-wrap:anywhere; }
  table.nob > tbody > tr > td, table.nob > tr > td { border:none; }
  .img { display:inline-block; border:1px dashed #9aa; color:#677; font-size:9pt; padding:8px 14px; }
  .h1 { font-size:16pt; font-weight:bold; } .h2 { font-size:13pt; font-weight:bold; } .h3 { font-size:12pt; font-weight:bold; }
</style></head>
<body><div class="page">
${headerHtml ? `<div class="hdr">${headerHtml}</div>` : ''}
${body ? blocks(body) : ''}
${footerHtml ? `<div class="ftr">${footerHtml}</div>` : ''}
</div></body></html>`;
}

/** Page size and margins of the section in CSS pixels (96 dpi), so landscape documents render landscape. */
function pageBox(sectPr: Element | undefined) {
  const px = (twips: number) => Math.round((twips / 1440) * 96);
  const pgSz = sectPr ? children(sectPr, 'pgSz')[0] : undefined;
  const pgMar = sectPr ? children(sectPr, 'pgMar')[0] : undefined;
  const n = (el: Element | undefined, a: string, d: number) => Number(el?.getAttributeNS(NS.w, a) || d);
  return {
    width: px(n(pgSz, 'w', 11906)),
    height: px(n(pgSz, 'h', 16838)),
    top: Math.min(px(n(pgMar, 'top', 1134)), 120),
    right: px(n(pgMar, 'right', 1134)),
    bottom: Math.min(px(n(pgMar, 'bottom', 1134)), 120),
    left: px(n(pgMar, 'left', 1134)),
  };
}

/** Width of the browser window for an HTML preview: the page plus a small grey frame. */
export function previewWindowWidth(html: string): number {
  const m = /data-page-width="(\d+)"/.exec(html);
  return (m ? Number(m[1]) : 794) + 106;
}

async function renderReferenced(
  pkg: WordPackage,
  sectPr: Element,
  kind: 'headerReference' | 'footerReference',
  rels: Map<string, string>,
): Promise<string> {
  const refs = children(sectPr, kind);
  const ref = refs.find((r) => r.getAttributeNS(NS.w, 'type') === 'default') ?? refs[0];
  const target = ref ? rels.get(ref.getAttributeNS(NS.r, 'id') ?? '') : undefined;
  if (!target) return '';
  const dom = await pkg.dom(`word/${target.replace(/^\.?\//, '')}`);
  return dom ? blocks(dom.documentElement as Element) : '';
}

function blocks(container: Element): string {
  let out = '';
  for (const c of children(container)) {
    switch (localName(c)) {
      case 'p':
        out += paragraph(c);
        break;
      case 'tbl':
        out += table(c);
        break;
      case 'sdt':
      case 'sdtContent':
      case 'customXml':
        out += blocks(localName(c) === 'sdt' ? (children(c, 'sdtContent')[0] ?? c) : c);
        break;
      default:
        break;
    }
  }
  return out;
}

function paragraph(p: Element): string {
  const pPr = children(p, 'pPr')[0];
  const style = pPr ? val(children(pPr, 'pStyle')[0]) : '';
  const jc = pPr ? val(children(pPr, 'jc')[0]) : '';
  const align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' ? 'justify' : '';
  const heading = /heading\s*1|titolo\s*1|^title$/i.test(style)
    ? 'h1'
    : /heading\s*2|titolo\s*2/i.test(style)
      ? 'h2'
      : /heading|titolo/i.test(style)
        ? 'h3'
        : '';
  const attrs = [heading ? `class="${heading}"` : '', align ? `style="text-align:${align}"` : '']
    .filter(Boolean)
    .join(' ');
  return `<p${attrs ? ` ${attrs}` : ''}>${inline(p)}</p>\n`;
}

function inline(container: Element): string {
  let out = '';
  for (const c of children(container)) {
    switch (localName(c)) {
      case 'r':
        out += run(c);
        break;
      case 'hyperlink':
      case 'smartTag':
      case 'ins':
      case 'fldSimple':
        out += inline(c);
        break;
      case 'sdt':
        out += inline(children(c, 'sdtContent')[0] ?? c);
        break;
      default:
        break;
    }
  }
  return out;
}

function run(r: Element): string {
  const rPr = children(r, 'rPr')[0];
  let text = '';
  for (const c of children(r)) {
    const n = localName(c);
    if (n === 't') text += esc(c.textContent ?? '');
    else if (n === 'br' || n === 'cr') text += '<br>';
    else if (n === 'tab') text += '&emsp;';
    else if (n === 'drawing' || n === 'pict') text += '<span class="img">image</span>';
  }
  if (!text || !rPr) return text;
  const css: string[] = [];
  if (on(rPr, 'b')) css.push('font-weight:bold');
  if (on(rPr, 'i')) css.push('font-style:italic');
  if (children(rPr, 'u')[0] && val(children(rPr, 'u')[0]) !== 'none') css.push('text-decoration:underline');
  const sz = val(children(rPr, 'sz')[0]);
  if (sz) css.push(`font-size:${Number(sz) / 2}pt`);
  const color = val(children(rPr, 'color')[0]);
  if (color && color !== 'auto') css.push(`color:#${color}`);
  if (on(rPr, 'caps')) css.push('text-transform:uppercase');
  if (on(rPr, 'vanish')) return '';
  return css.length ? `<span style="${css.join(';')}">${text}</span>` : text;
}

function table(tbl: Element): string {
  // Column widths from the grid (twips), as percentages, so wide tables fit the page.
  const grid = children(tbl, 'tblGrid')[0];
  const widths = grid ? children(grid, 'gridCol').map((g) => Number(g.getAttributeNS(NS.w, 'w') || 0)) : [];
  const total = widths.reduce((a, b) => a + b, 0);
  const tblPr = children(tbl, 'tblPr')[0];
  const tblBorders = tblPr ? children(tblPr, 'tblBorders')[0] : undefined;
  const borderless = tblBorders
    ? children(tblBorders).every((b) => ['nil', 'none'].includes(b.getAttributeNS(NS.w, 'val') ?? ''))
    : false;
  let out = `<table style="table-layout:fixed"${borderless ? ' class="nob"' : ''}>`;
  if (total > 0) {
    out += `<colgroup>${widths.map((w) => `<col style="width:${((w / total) * 100).toFixed(2)}%">`).join('')}</colgroup>`;
  }
  const rows = (container: Element): Element[] =>
    children(container).flatMap((c) =>
      localName(c) === 'tr' ? [c] : localName(c) === 'sdt' ? rows(children(c, 'sdtContent')[0] ?? c) : [],
    );
  for (const tr of rows(tbl)) {
    out += '<tr>';
    const cells = (container: Element): Element[] =>
      children(container).flatMap((c) =>
        localName(c) === 'tc' ? [c] : localName(c) === 'sdt' ? cells(children(c, 'sdtContent')[0] ?? c) : [],
      );
    for (const tc of cells(tr)) {
      const tcPr = children(tc, 'tcPr')[0];
      const span = tcPr ? Number(val(children(tcPr, 'gridSpan')[0]) || 1) : 1;
      const shade = tcPr ? children(tcPr, 'shd')[0]?.getAttributeNS(NS.w, 'fill') : '';
      const bg = shade && shade !== 'auto' ? ` style="background:#${shade}"` : '';
      out += `<td${span > 1 ? ` colspan="${span}"` : ''}${bg}>${blocks(tc)}</td>`;
    }
    out += '</tr>';
  }
  return `${out}</table>\n`;
}

function val(el: Element | undefined): string {
  return el?.getAttributeNS(NS.w, 'val') ?? '';
}

function on(rPr: Element, name: string): boolean {
  const el = children(rPr, name)[0];
  if (!el) return false;
  const v = val(el);
  return v === '' || v === '1' || v === 'true' || v === 'on';
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
