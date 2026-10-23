import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import { assertAllowedPath, readTemplateBytes } from '../docx/package.js';
import {
  browserMeasureHeight,
  browserScreenshot,
  browserToPdf,
  findBrowser,
  findLibreOffice,
  libreOfficeToPdf,
  outputBase,
  type PdfEngine,
  wordAvailable,
  wordToPdf,
} from './engines.js';
import { type FillResult, fillTemplate } from './fill.js';
import { previewWindowWidth, renderHtml } from './html.js';

export type PreviewFormat = 'docx' | 'html' | 'pdf' | 'png';

export interface PreviewOptions {
  /** Folder for the output files. Default: next to the template. */
  outDir?: string;
  language?: string;
  formats?: PreviewFormat[];
  /** Force a PDF engine; default tries LibreOffice, then Word, then a browser. */
  engine?: PdfEngine | 'auto';
}

export interface PreviewResult {
  files: Partial<Record<PreviewFormat, string>>;
  /** PNG slices of the whole document, page 1 first (files.png is the first). */
  pages: string[];
  pdfEngine: PdfEngine | null;
  language: string;
  stats: FillResult['stats'];
  issues: string[];
}

export async function renderPreview(templatePath: string, opts: PreviewOptions = {}): Promise<PreviewResult> {
  const source = resolve(templatePath);
  const formats = new Set<PreviewFormat>(opts.formats?.length ? opts.formats : ['docx', 'html', 'pdf']);
  const base = `${outputBase(source, opts.outDir ? resolve(opts.outDir) : undefined)}.preview`;
  assertAllowedPath(base);
  await mkdir(dirname(base), { recursive: true });

  const filled = await fillTemplate(await readTemplateBytes(source), { language: opts.language });
  const issues = [...filled.issues];
  const files: PreviewResult['files'] = {};

  // The .docx is always written: every PDF engine except the browser converts from it.
  const docx = `${base}.docx`;
  await writeFile(docx, await filled.pkg.toBuffer());
  if (formats.has('docx')) files.docx = docx;

  const needHtml = formats.has('html') || formats.has('png');
  const html = `${base}.html`;
  if (needHtml || formats.has('pdf')) {
    await writeFile(html, await renderHtml(filled.pkg, `Preview — ${basename(source)}`), 'utf8');
    if (formats.has('html')) files.html = html;
  }

  let pdfEngine: PdfEngine | null = null;
  if (formats.has('pdf')) {
    const pdf = `${base}.pdf`;
    const engine = opts.engine ?? 'auto';
    const attempts: Array<[PdfEngine, () => Promise<string | null>]> = [];
    const soffice = findLibreOffice();
    const browser = findBrowser();
    if ((engine === 'auto' || engine === 'libreoffice') && soffice) {
      attempts.push(['libreoffice', () => libreOfficeToPdf(soffice, docx, pdf)]);
    }
    if ((engine === 'auto' || engine === 'word') && (await wordAvailable())) {
      attempts.push(['word', () => wordToPdf(docx, pdf)]);
    }
    if ((engine === 'auto' || engine === 'browser') && browser) {
      attempts.push(['browser', () => browserToPdf(browser, html, pdf)]);
    }
    for (const [name, attempt] of attempts) {
      const error = await attempt();
      if (!error) {
        pdfEngine = name;
        files.pdf = pdf;
        break;
      }
      issues.push(`${name}: ${error}`);
    }
    if (!pdfEngine) {
      issues.push(
        attempts.length
          ? 'No PDF produced; open the .docx or .html preview instead.'
          : 'No PDF engine found. Install LibreOffice (any OS) or Microsoft Word (Windows), or use the .html preview.',
      );
    }
    if (pdfEngine === 'browser') issues.push('PDF made from the HTML preview: layout is simplified compared to Word.');
  }

  const pages: string[] = [];
  if (formats.has('png')) {
    const browser = findBrowser();
    if (!browser) issues.push('PNG: no Edge/Chrome/Chromium found');
    else {
      // One screenshot per page-sized slice: a single very tall image is unreadable once scaled down.
      const source = await readFile(html, 'utf8');
      const width = previewWindowWidth(source);
      const height = (await browserMeasureHeight(browser, html, width)) ?? (await estimateHeight(html));
      const count = Math.min(MAX_PNG_PAGES, Math.ceil(height / PNG_PAGE));
      for (let i = 0; i < count; i++) {
        const png = i === 0 ? `${base}.png` : `${base}.p${i + 1}.png`;
        let page = html;
        if (i > 0) {
          page = `${base}.p${i + 1}.tmp.html`;
          const shift = `<style>body{position:relative;top:-${i * PNG_PAGE}px}</style></head>`;
          await writeFile(page, source.replace('</head>', shift), 'utf8');
        }
        const error = await browserScreenshot(browser, page, png, PNG_PAGE, width);
        if (i > 0) await rm(page, { force: true });
        if (error) {
          issues.push(`PNG page ${i + 1}: ${error}`);
          break;
        }
        pages.push(png);
      }
      if (pages[0]) files.png = pages[0];
      if (Math.ceil(height / PNG_PAGE) > MAX_PNG_PAGES)
        issues.push(`PNG: only the first ${MAX_PNG_PAGES} pages; see the PDF/HTML for the rest.`);
    }
  }

  return { files, pages, pdfEngine, language: filled.language, stats: filled.stats, issues };
}

/**
 * Rough pixel height of the HTML preview, so the PNG slices cover the whole document
 * (a lines table is often on page 2). Fallback when the browser cannot measure it. Each paragraph or table row is at least one line;
 * long text wraps at ~95 characters, or ~40 inside table cells.
 */
async function estimateHeight(htmlPath: string): Promise<number> {
  const html = await readFile(htmlPath, 'utf8');
  let lines = 0;
  for (const block of html.match(/<(p|tr)[ >][\s\S]*?<\/(p|tr)>/g) ?? []) {
    const text = block.replace(/<br>/g, '\n').replace(/<[^>]+>/g, '');
    const perLine = block.startsWith('<tr') ? 40 : 95;
    lines += text.split('\n').reduce((n, l) => n + Math.max(1, Math.ceil(l.length / perLine)), 0);
  }
  return Math.max(PNG_PAGE, 160 + lines * 21);
}

const PNG_PAGE = 1250;
const MAX_PNG_PAGES = 8;

export function renderPreviewReport(r: PreviewResult): string {
  const s = r.stats;
  const lines = [
    `Preview (${r.language}${r.pdfEngine ? `, PDF via ${r.pdfEngine}` : ''}): ${s.fields} fields, ${s.labels} labels, ` +
      `${s.lists} list(s) with ${s.listItems} row(s), ${s.ifsTrue} conditions shown / ${s.ifsFalse} hidden.`,
    'Files:',
    ...Object.entries(r.files).map(([k, v]) => `- ${k}: ${v}`),
    ...(r.pages.length > 1 ? [`- png pages: ${r.pages.length} (${r.pages.slice(1).join(', ')})`] : []),
  ];
  if (r.issues.length) lines.push('Notes:', ...r.issues.map((i) => `- ${i}`));
  lines.push(
    'This is an approximation built from the sample data embedded in the template; the final check is the Docentric preview in D365 or the Designer.',
  );
  return lines.join('\n');
}
