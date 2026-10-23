import { deflateSync } from 'node:zlib';

/**
 * WordprocessingML fragments as strings, shaped like what Docentric Designer 6.x
 * writes: every element is a content control titled with the element type,
 * tagged `<projectId>-<elementId>`, coloured per type. Fields and Labels inside
 * tables wrap the whole cell; Lists that repeat table rows wrap the rows.
 */

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Designer colours per element type (from real templates). Field/SummaryField have none. */
export const ELEMENT_COLORS: Record<string, string> = {
  Label: 'FFCC00',
  If: '0000FF',
  List: 'FF00FF',
  Var: '008000',
  Group: 'FF00FF',
};

export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Points. */
  size?: number;
  /** Hex RRGGBB without '#'. */
  color?: string;
  font?: string;
  caps?: boolean;
}

export interface ParagraphStyle {
  align?: 'left' | 'center' | 'right' | 'justify';
  /** Style id, e.g. Heading1. */
  styleId?: string;
  /** Points of space after. */
  spaceAfter?: number;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
}

export function rPr(s: RunStyle | undefined): string {
  if (!s) return '';
  let x = '';
  if (s.font) x += `<w:rFonts w:ascii="${esc(s.font)}" w:hAnsi="${esc(s.font)}" w:cs="${esc(s.font)}"/>`;
  if (s.bold) x += '<w:b/><w:bCs/>';
  if (s.italic) x += '<w:i/><w:iCs/>';
  if (s.caps) x += '<w:caps/>';
  if (s.color) x += `<w:color w:val="${esc(s.color.replace('#', ''))}"/>`;
  if (s.size) x += `<w:sz w:val="${Math.round(s.size * 2)}"/><w:szCs w:val="${Math.round(s.size * 2)}"/>`;
  if (s.underline) x += '<w:u w:val="single"/>';
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}

export function pPr(s: ParagraphStyle | undefined): string {
  if (!s) return '';
  let x = '';
  if (s.styleId) x += `<w:pStyle w:val="${esc(s.styleId)}"/>`;
  if (s.keepNext) x += '<w:keepNext/>';
  if (s.pageBreakBefore) x += '<w:pageBreakBefore/>';
  if (s.spaceAfter !== undefined) x += `<w:spacing w:after="${Math.round(s.spaceAfter * 20)}"/>`;
  if (s.align) {
    const jc = s.align === 'justify' ? 'both' : s.align;
    x += `<w:jc w:val="${jc}"/>`;
  }
  return x ? `<w:pPr>${x}</w:pPr>` : '';
}

/** A run of text; newlines become line breaks. */
export function run(text: string, style?: RunStyle): string {
  const parts = text.split('\n').map((t) => `<w:t xml:space="preserve">${esc(t)}</w:t>`);
  return `<w:r>${rPr(style)}${parts.join('<w:br/>')}</w:r>`;
}

export function paragraph(inner: string, style?: ParagraphStyle, runStyle?: RunStyle): string {
  const props = pPr(style);
  // Paragraph mark formatting keeps empty paragraphs at the right height.
  const mark = runStyle ? rPr(runStyle).replace('<w:rPr>', '').replace('</w:rPr>', '') : '';
  const withMark = mark
    ? props
      ? props.replace('</w:pPr>', `<w:rPr>${mark}</w:rPr></w:pPr>`)
      : `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>`
    : props;
  return `<w:p>${withMark}${inner}</w:p>`;
}

export interface SdtOptions {
  alias: string;
  tag: string;
  /** Unique signed 32-bit integer (w:id). */
  id: number;
  runStyle?: RunStyle;
  showingPlaceholder?: boolean;
}

export function sdt(o: SdtOptions, content: string): string {
  const color = ELEMENT_COLORS[o.alias];
  return (
    `<w:sdt><w:sdtPr>${rPr(o.runStyle)}<w:alias w:val="${esc(o.alias)}"/><w:tag w:val="${esc(o.tag)}"/>` +
    `<w:id w:val="${o.id}"/>${o.showingPlaceholder ? '<w:showingPlcHdr/>' : ''}` +
    `${color ? `<w15:color w:val="${color}"/>` : ''}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

export interface CellOptions {
  /** Twips. */
  width?: number;
  shade?: string;
  span?: number;
  borders?: boolean;
  vAlign?: 'top' | 'center' | 'bottom';
}

export function cell(content: string, o: CellOptions = {}): string {
  let props = '';
  if (o.width) props += `<w:tcW w:w="${Math.round(o.width)}" w:type="dxa"/>`;
  if (o.span && o.span > 1) props += `<w:gridSpan w:val="${o.span}"/>`;
  if (o.borders === false) {
    props +=
      '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';
  }
  if (o.shade) props += `<w:shd w:val="clear" w:color="auto" w:fill="${esc(o.shade.replace('#', ''))}"/>`;
  if (o.vAlign) props += `<w:vAlign w:val="${o.vAlign}"/>`;
  // A cell must end with a paragraph.
  const body =
    /<w:p[ >]|<w:p\/>/.test(content) && !content.trimEnd().endsWith('</w:tbl>') ? content : `${content}<w:p/>`;
  return `<w:tc>${props ? `<w:tcPr>${props}</w:tcPr>` : ''}${body}</w:tc>`;
}

export function row(cells: string, o: { header?: boolean; cantSplit?: boolean } = {}): string {
  let props = '';
  if (o.cantSplit) props += '<w:cantSplit/>';
  if (o.header) props += '<w:tblHeader/>';
  return `<w:tr>${props ? `<w:trPr>${props}</w:trPr>` : ''}${cells}</w:tr>`;
}

export function table(rows: string, widths: number[], o: { borders?: boolean; layoutFixed?: boolean } = {}): string {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = (side: string) =>
    o.borders === false
      ? `<w:${side} w:val="nil"/>`
      : `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>`;
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${Math.round(total)}" w:type="dxa"/>` +
    `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` +
    `${o.layoutFixed ? '<w:tblLayout w:type="fixed"/>' : ''}` +
    '<w:tblCellMar><w:left w:w="85" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar>' +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
    `</w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${Math.round(w)}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`
  );
}

/** Inline picture placeholder, for Image elements (the engine swaps the bytes at print time). */
export function inlinePicture(relId: string, widthMm: number, heightMm: number, docPrId: number, name: string): string {
  const emu = (mm: number) => Math.round(mm * 36000);
  const cx = emu(widthMm);
  const cy = emu(heightMm);
  return (
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="${esc(name)}"/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  );
}

/** Solid-colour PNG, used as the design-time image placeholder (the engine swaps the bytes). */
export function solidPng(width: number, height: number, rgb: [number, number, number]): Buffer {
  const chunk = (type: string, data: Buffer) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body) >>> 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // RGB
  const rowBytes = Buffer.alloc(1 + width * 3);
  for (let x = 0; x < width; x++) rowBytes.set(rgb, 1 + x * 3);
  const raw = Buffer.concat(Array.from({ length: height }, () => rowBytes));
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** CRC-32 (PNG chunks). zlib.crc32 exists only from Node 22.2; this keeps Node 18+ working. */
function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (const byte of data) {
    c ^= byte;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  return (c ^ 0xffffffff) >>> 0;
}

/** mm → twips */
export const mmToTwips = (mm: number) => Math.round((mm / 25.4) * 1440);
