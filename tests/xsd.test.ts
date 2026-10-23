import { describe, expect, it } from 'vitest';
import { buildSchemaTree, countSchema } from '../src/schema/xsd.js';
import { parseXml } from '../src/xml/dom.js';

const schema = (inner: string) =>
  parseXml(`<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">${inner}</xs:schema>`, 'test.xsd');

describe('buildSchemaTree', () => {
  it('resolves named types, inline types, choices and repeated elements', () => {
    const root = buildSchemaTree(
      schema(`
        <xs:complexType name="Enum"><xs:attribute name="Value" type="xs:integer"/></xs:complexType>
        <xs:element name="ReportData"><xs:complexType>
          <xs:choice minOccurs="0" maxOccurs="unbounded">
            <xs:element name="Header"><xs:complexType>
              <xs:sequence><xs:element name="Level" type="Enum"/></xs:sequence>
              <xs:attribute name="Id" type="xs:string"/>
            </xs:complexType></xs:element>
            <xs:element name="Lines" maxOccurs="unbounded"><xs:complexType>
              <xs:attribute name="Qty" type="xs:decimal"/>
            </xs:complexType></xs:element>
          </xs:choice>
          <xs:attribute name="IsWithAmount" type="xs:boolean"/>
        </xs:complexType></xs:element>`),
    );
    expect(root?.name).toBe('ReportData');
    expect([...(root?.attributes.keys() ?? [])]).toEqual(['IsWithAmount']);
    expect(root?.children.get('Header')?.children.get('Level')?.attributes.get('Value')).toBe('integer');
    expect(root?.children.get('Lines')?.repeated).toBe(true);
    expect(root ? countSchema(root) : null).toEqual({ elements: 4, attributes: 4 });
  });

  it('merges complexContent extensions and follows element refs', () => {
    const root = buildSchemaTree(
      schema(`
        <xs:complexType name="Base"><xs:attribute name="A" type="xs:string"/></xs:complexType>
        <xs:complexType name="Derived"><xs:complexContent><xs:extension base="Base">
          <xs:attribute name="B" type="xs:string"/>
        </xs:extension></xs:complexContent></xs:complexType>
        <xs:element name="Root"><xs:complexType><xs:sequence>
          <xs:element name="Item" type="Derived"/>
          <xs:element ref="Shared"/>
        </xs:sequence></xs:complexType></xs:element>
        <xs:element name="Shared"><xs:complexType><xs:attribute name="S" type="xs:string"/></xs:complexType></xs:element>`),
    );
    expect([...(root?.children.get('Item')?.attributes.keys() ?? [])]).toEqual(['A', 'B']);
    expect(root?.children.get('Shared')?.attributes.has('S')).toBe(true);
  });

  it('stops on recursive types', () => {
    const root = buildSchemaTree(
      schema(`
        <xs:complexType name="Node"><xs:sequence><xs:element name="Node" type="Node"/></xs:sequence></xs:complexType>
        <xs:element name="Tree" type="Node"/>`),
    );
    expect(root?.children.get('Node')).toBeDefined();
  });

  it('returns null without a top-level element', () => {
    expect(buildSchemaTree(schema('<xs:complexType name="Only"/>'))).toBeNull();
  });
});
