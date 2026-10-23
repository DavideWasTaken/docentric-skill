import { esc } from './ooxml.js';

/** An element to write into the ProjectData registry. */
export interface NewElement {
  id: string;
  type: string;
  name?: string;
  source?: string | null;
  path?: string;
  resultType?: string;
  aggregatePath?: string;
  labelCode?: string;
  attributes?: Record<string, string>;
  children: string[];
}

export interface LabelDef {
  code: string;
  texts: Record<string, string>;
  /** Template labels are 'Custom'; cached D365/ISV texts have no section. */
  custom: boolean;
}

export interface ProjectDataInput {
  productVersion: string;
  ddspUri: string;
  report: { id: string; group: string; dspClass: string; dataSourceType: string };
  languages: string[];
  labels: LabelDef[];
  /** A complete <DataSources …>…</DataSources> element. */
  dataSourcesXml: string;
  elements: NewElement[];
}

export const TEMPLATES_NS = 'http://schemas.docentric.com/dynamics-ax/templates';
export const DEFAULT_PRODUCT_VERSION = '6.1.62.1715';
const ZERO_GUID = '{00000000-0000-0000-0000-000000000000}';

/** Result types the Designer writes per element type. */
export const RESULT_TYPES: Record<string, string> = {
  Field: 'System.String',
  If: 'System.Boolean',
  Var: 'System.Object',
  Image: 'System.Byte[]',
};

export function elementXml(e: NewElement): string {
  const attrs = Object.entries(e.attributes ?? {})
    .map(([k, v]) => ` ${k}="${esc(v)}"`)
    .join('');
  const name = e.name ? ` Name="${esc(e.name)}"` : '';
  const kids = e.children.length
    ? `<ChildElements>${e.children.map((id) => `<ElementRef Id="${id}"/>`).join('')}</ChildElements>`
    : '';
  const value = e.labelCode ? `<Value LabelCode="${esc(e.labelCode)}" IsCodeManuallyEntered="1"/>` : '';
  const result = e.resultType ?? RESULT_TYPES[e.type];
  const binding =
    e.path !== undefined
      ? `<ValueBinding DataKind="2"${e.source ? ` Source="${esc(e.source)}"` : ''} Path="${esc(e.path)}"${result ? ` ResultDataType="${result}"` : ''}/>`
      : '';
  const sort = e.type === 'List' || e.type === 'SummaryField' ? '<SortDescriptors/>' : '';
  const aggregate =
    e.aggregatePath !== undefined
      ? `<AggregateBinding DataKind="2" Path="${esc(e.aggregatePath)}" ResultDataType="System.Double"/>`
      : '';
  return `<Element Type="${esc(e.type)}" Id="${e.id}"${name}${attrs}>${kids}${value}${binding}${sort}${aggregate}</Element>`;
}

export function labelXml(l: LabelDef, prefix = 'd2p1'): string {
  const texts = Object.entries(l.texts)
    .map(([lang, v]) => `<text ${prefix}:language="${esc(lang)}" ${prefix}:value="${esc(v)}"/>`)
    .join('');
  if (!l.custom) return `<label ${prefix}:code="${esc(l.code)}">${texts}</label>`;
  const name = l.code.split(':').pop() ?? l.code;
  return `<label ${prefix}:name="${esc(name)}" ${prefix}:code="${esc(l.code)}" ${prefix}:section="Custom">${texts}</label>`;
}

export function projectDataXml(p: ProjectDataInput): string {
  const r = p.report;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' +
    `<ProjectData ProductId="3" ProductVersion="${esc(p.productVersion)}" xmlns="http://Docentric.org/ProjectData.xml">` +
    '<UISettings ELEMENT_APPEAREANCE_TYPE="1" xmlns=""><ElementsPane IsVisible="1" DockPosition="2" Width="510" Height="862"/></UISettings>' +
    `<dataSourcePackage d2p1:uri="${esc(p.ddspUri)}" xmlns:d2p1="${TEMPLATES_NS}" xmlns="${TEMPLATES_NS}">` +
    `<ssrsReportInfo d2p1:reportId="${esc(r.id)}" d2p1:reportGroupId="${esc(r.group)}" ` +
    `d2p1:dataSourceProviderClassName="${esc(r.dspClass)}" d2p1:erModelMappingGuid="${ZERO_GUID}" ` +
    `d2p1:erDataModelGuid="${ZERO_GUID}" d2p1:erSolutionGuid="${ZERO_GUID}" d2p1:dataSourceType="${esc(r.dataSourceType)}"/>` +
    `<languages d2p1:defaultLanguage="${esc(p.languages[0] ?? 'en-us')}">${p.languages.map((l) => `<language d2p1:value="${esc(l)}"/>`).join('')}</languages>` +
    `<labels>${p.labels.map((l) => labelXml(l)).join('')}</labels>` +
    '</dataSourcePackage>' +
    p.dataSourcesXml +
    `<Elements xmlns="">${p.elements.map(elementXml).join('')}</Elements>` +
    '</ProjectData>'
  );
}

/** Random ids in the ranges the Designer uses: positive uint32 element ids, signed int32 w:id. */
export class IdAllocator {
  private readonly used = new Set<string>();

  constructor(existing: Iterable<string> = []) {
    for (const id of existing) this.used.add(id);
  }

  elementId(): string {
    for (;;) {
      const id = String(Math.floor(Math.random() * 4_200_000_000) + 10_000_000);
      if (!this.used.has(id)) {
        this.used.add(id);
        return id;
      }
    }
  }

  static sdtId(): number {
    return Math.floor(Math.random() * 4_294_967_295) - 2_147_483_648;
  }

  static projectId(): string {
    return String(Math.floor(Math.random() * 90_000_000) + 10_000_000);
  }
}
