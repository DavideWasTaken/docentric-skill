import type { AttrType, DataSourceShape, ElementShape } from '../create/spec.js';
import { esc } from './ooxml.js';

/**
 * Builds the XSD and sample XML that a data source carries inside the template,
 * from a declared shape. Used when there is no .ddsp yet (new template, SSRS
 * conversion): the shape comes from the agent — typically from the DSP temp
 * tables read with d365fo-mcp — and is marked provisional until a real DDSP is
 * imported.
 */

const XS_TYPE: Record<AttrType, string> = {
  string: 'xs:string',
  decimal: 'xs:decimal',
  int: 'xs:int',
  date: 'xs:date',
  dateTime: 'xs:dateTime',
  boolean: 'xs:boolean',
  base64Binary: 'xs:base64Binary',
};

const ENUM_TYPES =
  '<xs:complexType name="Enum"><xs:attribute name="Value" type="xs:integer" use="optional"/>' +
  '<xs:attribute name="Name" type="xs:string" use="optional"/><xs:attribute name="Text" type="xs:string" use="optional"/></xs:complexType>' +
  '<xs:complexType name="UserDefEnum"><xs:attribute name="Name" type="xs:string" use="optional"/>' +
  '<xs:attribute name="Label" type="xs:string" use="optional"/></xs:complexType>';

export function shapeToXsd(shape: DataSourceShape): string {
  return `${ENUM_TYPES}${elementXsd(shape.root, shape, false)}`;
}

function elementXsd(name: string, shape: ElementShape, repeated: boolean): string {
  const kids = Object.entries(shape.elements ?? {})
    .map(([n, s]) => elementXsd(n, s, !!s.repeated))
    .join('');
  const attrs = Object.entries(shape.attributes ?? {})
    .map(([n, t]) => `<xs:attribute name="${esc(n)}" type="${XS_TYPE[t]}" use="optional"/>`)
    .join('');
  const choice = kids ? `<xs:choice minOccurs="0" maxOccurs="unbounded">${kids}</xs:choice>` : '';
  return `<xs:element${repeated ? ' maxOccurs="unbounded"' : ''} name="${esc(name)}"><xs:complexType>${choice}${attrs}</xs:complexType></xs:element>`;
}

/** Plausible sample values so the preview shows something readable: 3 rows per repeated element. */
export function shapeToSample(shape: DataSourceShape): string {
  return shape.sample ?? sampleElement(shape.root, shape, 0);
}

function sampleElement(name: string, shape: ElementShape, index: number): string {
  const attrs = Object.entries(shape.attributes ?? {})
    .map(([n, t]) => ` ${esc(n)}="${esc(sampleValue(n, t, index))}"`)
    .join('');
  const kids = Object.entries(shape.elements ?? {})
    .map(([n, s]) => Array.from({ length: s.repeated ? 3 : 1 }, (_, i) => sampleElement(n, s, i)).join(''))
    .join('');
  return kids ? `<${name}${attrs}>${kids}</${name}>` : `<${name}${attrs}/>`;
}

function sampleValue(name: string, type: AttrType, i: number): string {
  switch (type) {
    case 'decimal':
      return (10 * (i + 1) + 0.5 * (i + 1)).toFixed(2);
    case 'int':
      return String(i + 1);
    case 'date':
      return `2026-01-${String(15 + i).padStart(2, '0')}`;
    case 'dateTime':
      return `2026-01-${String(15 + i).padStart(2, '0')}T00:00:00`;
    case 'boolean':
      return 'true';
    case 'base64Binary':
      return '';
    default:
      return `${name} ${i + 1}`;
  }
}

export const SECTION_CODE: Record<string, string> = { MainData: '1', GeneralData: '2', Parameters: '3' };

export function dataSourceXml(name: string, shape: DataSourceShape, prefix: string): string {
  const section = SECTION_CODE[shape.section ?? (name in SECTION_CODE ? name : 'MainData')] ?? '1';
  return (
    `<DataSource Name="${esc(name)}" IsDefault="0" DataKind="2" ${prefix}:dataSectionType="${section}">` +
    `<Schema><XsdList><MainXsd xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:schema>${shapeToXsd(shape)}</xs:schema></MainXsd></XsdList>` +
    '<XmlNamespaceMappings><XmlNamespaceMapping Prefix="xs" Name="http://www.w3.org/2001/XMLSchema"/></XmlNamespaceMappings></Schema>' +
    `<PreviewData><SampleXml>${shapeToSample(shape)}</SampleXml></PreviewData></DataSource>`
  );
}
