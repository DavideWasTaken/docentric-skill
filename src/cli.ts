import { readFile, writeFile } from 'node:fs/promises';
import { dirname, extname, resolve } from 'node:path';
import { parse as parseYaml, stringify as toYaml } from 'yaml';
import { createTemplate } from './create/createTemplate.js';
import { loadTemplateFile } from './docx/loadTemplate.js';
import { assertAllowedPath } from './docx/package.js';
import { editTemplate } from './edit/editTemplate.js';
import { inspectTemplate } from './inspect/inspectTemplate.js';
import { findBrowser, findLibreOffice, wordAvailable } from './preview/engines.js';
import { type PreviewFormat, renderPreview, renderPreviewReport } from './preview/renderPreview.js';
import { convertSsrs } from './ssrs/convert.js';
import { renderFindings, type Severity, validateTemplate } from './validate/validateTemplate.js';
import { VERSION } from './version.js';

/**
 * The command line the skill drives. Output is plain text written for the agent
 * reading it (and for people); every command that writes says where.
 */

const HELP = `docentric ${VERSION} — Docentric AX templates (Dynamics 365 F&O), from the command line

Usage: node docentric.mjs <command> [arguments]

  inspect <template.docx> [--schema] [--labels] [--source MainData] [--depth 3]
      Report, DSP class, languages, data sources, element tree (ids, bindings, data context).
  validate <template.docx> [--json] [--min error|warning|info]
      Offline checks DCX001–DCX019. Exit code 1 when there are errors.
  preview <template.docx> [--out <dir>] [--formats docx,html,pdf,png] [--lang it] [--engine auto|libreoffice|word|browser]
      Fills the template with its sample data and writes <name>.preview.* files.
  create <spec.yaml|spec.json|-> <out.docx> [--preview]
      New template from a layout spec ('-' reads the spec from stdin).
  edit <template.docx> <ops.yaml|ops.json|-> [--out <file.docx>] [--dry-run] [--force] [--preview]
      Applies edit operations to a copy ('-' reads the operations from stdin).
  convert-ssrs <AxReport.xml> <out.docx> [--design <name>]
      SSRS report → <out>.spec.yaml + template.
  dump <template.docx>
      Stable text view, for git diff (textconv).
  doctor
      Checks Node.js and the optional PDF/PNG converters.

Environment: DOCENTRIC_ALLOWED_ROOTS (restrict file access), DOCENTRIC_SOFFICE, DOCENTRIC_BROWSER.
`;

export class UsageError extends Error {}

export async function runCli(argv: string[], stdin: () => Promise<string> = readStdin): Promise<number> {
  const [cmd, ...rest] = argv;
  const { positional, flags } = parseArgs(rest);
  const need = (n: number, usage: string) => {
    if (positional.length < n) throw new UsageError(`usage: docentric ${usage}`);
  };
  const structured = async (path: string) =>
    path === '-' ? parseStructured(await stdin(), '.yaml') : readStructured(path);

  switch (cmd) {
    case undefined:
    case '--help':
    case '-h':
    case 'help':
      console.log(HELP);
      return 0;
    case '--version':
    case '-v':
      console.log(VERSION);
      return 0;
    case 'inspect': {
      need(1, 'inspect <template.docx>');
      const include: Array<'summary' | 'elements' | 'schema' | 'labels'> = ['summary', 'elements'];
      if (flags.schema) include.push('schema');
      if (flags.labels) include.push('labels');
      console.log(
        inspectTemplate(await loadTemplateFile(positional[0] as string), {
          include,
          dataSource: str(flags.source),
          schemaDepth: flags.depth ? Number(flags.depth) : undefined,
        }),
      );
      return 0;
    }
    case 'validate': {
      need(1, 'validate <template.docx>');
      const r = validateTemplate(await loadTemplateFile(positional[0] as string));
      console.log(flags.json ? JSON.stringify(r, null, 2) : renderFindings(r, (str(flags.min) as Severity) ?? 'info'));
      return r.counts.error > 0 ? 1 : 0;
    }
    case 'preview': {
      need(1, 'preview <template.docx>');
      const r = await renderPreview(positional[0] as string, {
        outDir: str(flags.out),
        language: str(flags.lang),
        formats: str(flags.formats)?.split(',') as PreviewFormat[] | undefined,
        engine: str(flags.engine) as 'auto' | undefined,
      });
      console.log(renderPreviewReport(r));
      return 0;
    }
    case 'create': {
      need(2, 'create <spec.yaml|-> <out.docx>');
      const specArg = positional[0] as string;
      const baseDir = specArg === '-' ? process.cwd() : dirname(resolve(specArg));
      const r = await createTemplate(await structured(specArg), positional[1] as string, { baseDir });
      console.log(
        `Created ${r.path} — ${r.elements} elements, data sources: ${r.dataSources.join(', ')}` +
          (r.provisional.length ? ` (provisional: ${r.provisional.join(', ')})` : ''),
      );
      console.log(renderFindings(r.validation, 'warning'));
      if (flags.preview)
        console.log(renderPreviewReport(await renderPreview(r.path, { formats: ['pdf', 'png', 'html'] })));
      return r.validation.counts.error > 0 ? 1 : 0;
    }
    case 'edit': {
      need(2, 'edit <template.docx> <ops.yaml|->');
      const r = await editTemplate(positional[0] as string, await structured(positional[1] as string), {
        outputPath: str(flags.out),
        dryRun: !!flags['dry-run'],
        force: !!flags.force,
      });
      for (const a of r.applied) {
        const created = a.newElements.length ? ` → new element(s) ${a.newElements.map((i) => `#${i}`).join(', ')}` : '';
        console.log(`- ${a.op}: ${a.detail}${created}`);
      }
      const count = (v: typeof r.before) => `${v.counts.error} error(s), ${v.counts.warning} warning(s)`;
      console.log(`Validation before: ${count(r.before)} · after: ${count(r.after)}`);
      if (r.written) console.log(`Saved ${r.path}`);
      else if (flags['dry-run']) console.log('Dry run — nothing written.');
      else console.log('NOT saved: the edit introduces validation errors (use --force to save anyway).');
      for (const f of r.newErrors) console.log(`  ${f.code} ${f.message}${f.hint ? ` (${f.hint})` : ''}`);
      if (flags.preview && r.path)
        console.log(renderPreviewReport(await renderPreview(r.path, { formats: ['pdf', 'png', 'html'] })));
      return r.written || flags['dry-run'] ? 0 : 1;
    }
    case 'convert-ssrs': {
      need(2, 'convert-ssrs <AxReport.xml> <out.docx>');
      const conv = await convertSsrs(positional[0] as string, { design: str(flags.design) });
      const out = resolve(positional[1] as string);
      const specPath = `${out.replace(/\.docx$/i, '')}.spec.yaml`;
      assertAllowedPath(specPath);
      const created = await createTemplate(conv.spec, out, { baseDir: dirname(specPath) });
      await writeFile(specPath, toYaml(conv.spec), 'utf8');
      const s = conv.stats;
      console.log(
        `Converted: ${s.textboxes} textboxes, ${s.tables} table(s), ${s.images} image(s); ` +
          `${s.untranslated} expression(s) left as ⟦SSRS⟧ markers.`,
      );
      console.log(`Spec: ${specPath}`);
      console.log(
        `Template: ${created.path} (${created.elements} elements, provisional: ${created.provisional.join(', ')})`,
      );
      console.log(conv.notes.length ? `Notes:\n${conv.notes.map((n) => `- ${n}`).join('\n')}` : 'Notes: none');
      return 0;
    }
    case 'dump': {
      need(1, 'dump <template.docx>');
      console.log(
        inspectTemplate(await loadTemplateFile(positional[0] as string), {
          include: ['summary', 'elements', 'labels'],
        }),
      );
      return 0;
    }
    case 'doctor':
      return doctor();
    default:
      throw new UsageError(`unknown command '${cmd}'\n\n${HELP}`);
  }
}

async function doctor(): Promise<number> {
  const node = Number(process.versions.node.split('.')[0]);
  const soffice = findLibreOffice();
  const word = await wordAvailable();
  const browser = findBrowser();
  const line = (ok: boolean, label: string, detail: string) =>
    console.log(`${ok ? '✔' : '·'} ${label.padEnd(14)} ${detail}`);
  console.log(`docentric ${VERSION}\n`);
  line(node >= 18, 'Node.js', `${process.versions.node}${node >= 18 ? '' : ' — version 18 or later is required'}`);
  line(!!soffice, 'LibreOffice', soffice ?? 'not found (optional: PDF previews on any OS; set DOCENTRIC_SOFFICE)');
  line(word, 'Word', word ? 'available (PDF previews on Windows)' : 'not found (optional)');
  line(!!browser, 'Browser', browser ?? 'not found (optional: PNG previews; set DOCENTRIC_BROWSER)');
  line(true, 'Allowed roots', process.env.DOCENTRIC_ALLOWED_ROOTS ?? 'not set (any path can be read/written)');
  console.log(`\nPreview: ${soffice || word || browser ? 'PDF' : 'HTML only'}${browser ? ' + PNG' : ''}.`);
  return node >= 18 ? 0 : 1;
}

async function readStructured(path: string): Promise<unknown> {
  const full = resolve(path);
  assertAllowedPath(full);
  return parseStructured(await readFile(full, 'utf8'), extname(path));
}

function parseStructured(text: string, ext: string): unknown {
  // YAML is a superset of JSON, so .json files and JSON on stdin parse too; JSON.parse gives better errors.
  return ext.toLowerCase() === '.json' ? JSON.parse(text) : parseYaml(text);
}

async function readStdin(): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const c of process.stdin) chunks.push(c as Buffer);
  return Buffer.concat(chunks).toString('utf8');
}

function parseArgs(args: string[]): { positional: string[]; flags: Record<string, string | boolean> } {
  const positional: string[] = [];
  const flags: Record<string, string | boolean> = {};
  for (let i = 0; i < args.length; i++) {
    const a = args[i] as string;
    if (a.startsWith('--')) {
      const [k, v] = a.slice(2).split('=', 2) as [string, string | undefined];
      const next = args[i + 1];
      if (v !== undefined) flags[k] = v;
      else if (next !== undefined && !next.startsWith('--') && VALUE_FLAGS.has(k)) {
        flags[k] = next;
        i++;
      } else flags[k] = true;
    } else positional.push(a);
  }
  return { positional, flags };
}

const VALUE_FLAGS = new Set(['out', 'formats', 'lang', 'engine', 'min', 'source', 'depth', 'design']);

const str = (v: string | boolean | undefined) => (typeof v === 'string' ? v : undefined);
