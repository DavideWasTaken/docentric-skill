import { mkdir } from 'node:fs/promises';
import { dirname, isAbsolute, resolve } from 'node:path';
import { type Element, XMLSerializer } from '@xmldom/xmldom';
import { findProjectDataPart, loadTemplate } from '../docx/loadTemplate.js';
import { readTemplateBytes, WordPackage } from '../docx/package.js';
import { isTemplateLabel, TEMPLATE_LABEL_PREFIX } from '../model/projectData.js';
import type { DocentricTemplate } from '../model/types.js';
import { type ValidationResult, validateTemplate } from '../validate/validateTemplate.js';
import { assemble } from '../write/assemble.js';
import { dataSourceXml } from '../write/schemaXml.js';
import {
  DEFAULT_PRODUCT_VERSION,
  IdAllocator,
  type LabelDef,
  projectDataXml,
  TEMPLATES_NS,
} from '../write/projectDataXml.js';
import { child } from '../xml/dom.js';
import { IMAGE_REL_TOKEN, LayoutCompiler, pageGeometry, type Scope } from './compile.js';
import { type LayoutSpec, parseLayoutSpec, SpecError } from './spec.js';

export const PROVISIONAL_PROPERTY = 'DocentricSkill.ProvisionalDataSources';

export interface CreateResult {
  path: string;
  elements: number;
  dataSources: string[];
  provisional: string[];
  validation: ValidationResult;
}

export interface CreateOptions {
  /** Folder that relative paths in the spec (dataSources.from, baseDocument) resolve against. */
  baseDir?: string;
}

export async function createTemplate(
  specInput: unknown,
  outputPath: string,
  opts: CreateOptions = {},
): Promise<CreateResult> {
  const spec = parseLayoutSpec(specInput);
  const baseDir = opts.baseDir ?? process.cwd();
  const at = (p: string) => (isAbsolute(p) ? p : resolve(baseDir, p));

  const ds = await resolveDataSources(spec, at);
  const languages = spec.languages ?? ds.languages ?? ['en-us'];
  const defaultSource =
    spec.defaultSource ?? (ds.names.includes('MainData') ? 'MainData' : (ds.names[0] ?? 'MainData'));
  if (!ds.names.includes(defaultSource)) {
    throw new SpecError(`defaultSource '${defaultSource}' is not a data source (have: ${ds.names.join(', ')}).`);
  }

  const projectId = IdAllocator.projectId();
  const compiler = new LayoutCompiler(projectId, new IdAllocator(), ds.names);
  const page = pageGeometry(spec.page);
  const scope: Scope = { source: defaultSource, parentId: null, width: page.contentWidth };
  const header = spec.header ? compiler.blocks(spec.header, scope) : undefined;
  const body = compiler.blocks(spec.body, scope);
  const footer = spec.footer ? compiler.blocks(spec.footer, scope) : undefined;

  const labels = buildLabels(spec, ds.source, compiler.elements.map((e) => e.labelCode).filter(Boolean) as string[]);
  const pdXml = projectDataXml({
    productVersion: ds.source?.productVersion ?? DEFAULT_PRODUCT_VERSION,
    ddspUri: ds.source?.report.ddspUri ?? '',
    report: {
      id: spec.report?.id ?? ds.source?.report.reportId ?? '',
      group: spec.report?.group ?? ds.source?.report.reportGroupId ?? 'DEFAULT',
      dspClass: spec.report?.dspClass ?? ds.source?.report.dataSourceProviderClass ?? '',
      dataSourceType: ds.source?.report.dataSourceType ?? 'SSRS',
    },
    languages,
    labels,
    dataSourcesXml: ds.xml,
    elements: compiler.elements,
  });

  const base = spec.baseDocument ? await WordPackage.open(await readTemplateBytes(at(spec.baseDocument))) : undefined;
  const pkg = await assemble({
    body,
    header,
    footer,
    projectDataXml: pdXml,
    page,
    font: { name: spec.font?.name ?? 'Calibri', size: spec.font?.size ?? 10 },
    title: spec.title ?? spec.report?.id ?? 'Docentric template',
    imageToken: IMAGE_REL_TOKEN,
    base,
    customProperties: ds.provisional.length ? { [PROVISIONAL_PROPERTY]: ds.provisional.join(',') } : undefined,
  });

  const template = await loadTemplate(pkg);
  const validation = validateTemplate(template);
  const full = resolve(outputPath);
  await mkdir(dirname(full), { recursive: true });
  await pkg.saveAs(full);
  return {
    path: full,
    elements: compiler.elements.length,
    dataSources: ds.names,
    provisional: ds.provisional,
    validation,
  };
}

interface ResolvedSources {
  xml: string;
  names: string[];
  provisional: string[];
  languages?: string[];
  source?: DocentricTemplate;
}

async function resolveDataSources(spec: LayoutSpec, at: (p: string) => string): Promise<ResolvedSources> {
  if ('from' in spec.dataSources) {
    const pkg = await WordPackage.open(await readTemplateBytes(at(spec.dataSources.from)));
    const template = await loadTemplate(pkg);
    const part = (await findProjectDataPart(pkg)) as string;
    const pd = await pkg.dom(part);
    const container = pd ? child(pd.documentElement as Element, 'DataSources') : null;
    if (!container) throw new SpecError(`${spec.dataSources.from} has no data sources to copy.`);
    // Copied verbatim: schema, sample data and namespace mappings stay exactly as the DDSP made them.
    const xml = new XMLSerializer().serializeToString(container);
    return {
      xml,
      names: template.dataSources.map((d) => d.name),
      provisional: template.provisionalDataSources,
      languages: template.languages.length ? template.languages : undefined,
      source: template,
    };
  }
  const entries = Object.entries(spec.dataSources.schema);
  if (!entries.length) throw new SpecError('dataSources.schema declares no data source.');
  const xml = `<DataSources xmlns="" xmlns:ns1="${TEMPLATES_NS}">${entries.map(([name, shape]) => dataSourceXml(name, shape, 'ns1')).join('')}</DataSources>`;
  return { xml, names: entries.map(([n]) => n), provisional: entries.map(([n]) => n) };
}

function buildLabels(spec: LayoutSpec, source: DocentricTemplate | undefined, used: string[]): LabelDef[] {
  const out = new Map<string, LabelDef>();
  // Cached D365/ISV texts and the source template's own labels that the new layout uses.
  for (const l of source?.labels ?? []) {
    if (!isTemplateLabel(l)) out.set(l.code, { code: l.code, texts: l.texts, custom: false });
    else if (used.includes(l.code)) out.set(l.code, { code: l.code, texts: l.texts, custom: true });
  }
  for (const [name, texts] of Object.entries(spec.labels ?? {})) {
    const code = name.startsWith('@') ? name : `${TEMPLATE_LABEL_PREFIX}${name}`;
    out.set(code, { code, texts, custom: true });
  }
  return [...out.values()];
}
