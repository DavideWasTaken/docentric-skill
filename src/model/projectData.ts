import { buildSchemaTree } from '../schema/xsd.js';
import { allAttrs, attr, child, children, type Element, localName, path } from '../xml/dom.js';
import type {
  Binding,
  DataSectionType,
  DataSource,
  DocentricTemplate,
  ReportInfo,
  TemplateElement,
  TemplateLabel,
} from './types.js';

export const PROJECT_DATA_NS = 'http://Docentric.org/ProjectData.xml';

export const TEMPLATE_LABEL_PREFIX = '@DOC_LABEL:';

/**
 * The labels block mixes the template's own labels (@DOC_LABEL:…, section Custom)
 * with cached texts of the D365/ISV labels the schema uses (@SYS…, @DocentricAX:…,
 * no name or section). Only the former belong to the template.
 */
export function isTemplateLabel(l: TemplateLabel): boolean {
  return l.code.startsWith(TEMPLATE_LABEL_PREFIX) || l.section === 'Custom';
}

export function isProjectData(root: Element): boolean {
  return localName(root) === 'ProjectData';
}

/** Parses the ProjectData custom XML part. Content controls are filled in by the package loader. */
export function parseProjectData(
  root: Element,
  partName: string,
): Omit<DocentricTemplate, 'contentControls' | 'provisionalDataSources'> {
  const pkg = child(root, 'dataSourcePackage');
  const reportEl = pkg ? child(pkg, 'ssrsReportInfo') : null;
  const languagesEl = pkg ? child(pkg, 'languages') : null;

  const report: ReportInfo = {
    dataSourceType: reportEl ? attr(reportEl, 'dataSourceType') : null,
    reportId: reportEl ? attr(reportEl, 'reportId') : null,
    reportGroupId: reportEl ? attr(reportEl, 'reportGroupId') : null,
    dataSourceProviderClass: reportEl ? attr(reportEl, 'dataSourceProviderClassName') : null,
    ddspUri: pkg ? attr(pkg, 'uri') : null,
  };

  return {
    productVersion: attr(root, 'ProductVersion'),
    report,
    defaultLanguage: languagesEl ? attr(languagesEl, 'defaultLanguage') : null,
    languages: languagesEl
      ? children(languagesEl, 'language')
          .map((l) => attr(l, 'value'))
          .filter((v): v is string => !!v)
      : [],
    labels: pkg ? parseLabels(pkg) : [],
    dataSources: parseDataSources(root),
    elements: parseElements(root),
    projectDataPart: partName,
  };
}

function parseLabels(pkg: Element): TemplateLabel[] {
  const labelsEl = child(pkg, 'labels');
  if (!labelsEl) return [];
  return children(labelsEl, 'label').map((l) => {
    const texts: Record<string, string> = {};
    for (const t of children(l, 'text')) {
      const lang = attr(t, 'language');
      if (lang) texts[lang] = attr(t, 'value') ?? '';
    }
    return {
      name: attr(l, 'name') ?? '',
      code: attr(l, 'code') ?? '',
      section: attr(l, 'section') ?? '',
      texts,
    };
  });
}

const SECTION_BY_CODE: Record<string, DataSectionType> = {
  '1': 'MainData',
  '2': 'GeneralData',
  '3': 'Parameters',
};

function parseDataSources(root: Element): DataSource[] {
  const container = child(root, 'DataSources');
  if (!container) return [];
  return children(container, 'DataSource').map((ds) => {
    const xsd = path(ds, 'Schema', 'XsdList', 'MainXsd');
    const schema = xsd ? child(xsd, 'schema') : null;
    const sample = path(ds, 'PreviewData', 'SampleXml');
    return {
      name: attr(ds, 'Name') ?? '',
      section: SECTION_BY_CODE[attr(ds, 'dataSectionType') ?? ''] ?? 'Unknown',
      root: schema ? buildSchemaTree(schema) : null,
      hasSampleXml: !!sample && children(sample).length > 0,
    };
  });
}

const KNOWN_CHILDREN = new Set(['ValueBinding', 'AggregateBinding', 'ChildElements', 'Value', 'SortDescriptors']);

function parseElements(root: Element): Map<string, TemplateElement> {
  const out = new Map<string, TemplateElement>();
  const container = child(root, 'Elements');
  if (!container) return out;
  for (const el of children(container, 'Element')) {
    const id = attr(el, 'Id');
    if (!id) continue;
    const { Type, Id: _id, Name, ...rest } = allAttrs(el);
    const childEls = child(el, 'ChildElements');
    const value = child(el, 'Value');
    for (const c of children(el)) {
      if (!KNOWN_CHILDREN.has(localName(c))) rest[`child:${localName(c)}`] = 'present';
    }
    out.set(id, {
      id,
      type: Type ?? 'Unknown',
      name: Name ?? null,
      binding: parseBinding(child(el, 'ValueBinding')),
      aggregateBinding: parseBinding(child(el, 'AggregateBinding')),
      labelCode: value ? attr(value, 'LabelCode') : null,
      childIds: childEls
        ? children(childEls, 'ElementRef')
            .map((r) => attr(r, 'Id'))
            .filter((v): v is string => !!v)
        : [],
      attributes: rest,
    });
  }
  return out;
}

function parseBinding(el: Element | null): Binding | null {
  if (!el) return null;
  return {
    source: attr(el, 'Source'),
    path: attr(el, 'Path') ?? '',
    resultType: attr(el, 'ResultDataType'),
  };
}
