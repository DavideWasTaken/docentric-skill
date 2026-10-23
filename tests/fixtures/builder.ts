import JSZip from 'jszip';

/**
 * Builds synthetic Docentric templates in memory. Every fixture is generated —
 * no real customer template is ever committed (real ones embed sample data with
 * names, addresses and VAT numbers, and tenant sensitivity labels).
 *
 * The XML mirrors what Docentric Designer 6.x writes: a ProjectData custom XML
 * part with the data source package, data sources (XSD + sample) and element
 * registry, and one block-level content control per element tagged
 * `<projectId>-<elementId>` with the element type as its title.
 */

export interface FixtureElement {
  id: string;
  type: string;
  name?: string;
  source?: string;
  path?: string;
  aggregatePath?: string;
  labelCode?: string;
  attrs?: Record<string, string>;
  children?: FixtureElement[];
  /** Default true. False → registered in ProjectData but no content control. */
  placed?: boolean;
  /** Place in word/header1.xml instead of the body (top-level elements only). */
  inHeader?: boolean;
  /** Overrides the content control title (normally the element type). */
  alias?: string;
  /** Extra ElementRef ids written in ChildElements without a matching element. */
  danglingChildIds?: string[];
}

export interface FixtureDataSource {
  name: string;
  section: 1 | 2 | 3;
  /** Children of xs:schema (named complexTypes + the root xs:element). */
  schema: string;
  sample?: string;
}

export interface FixtureSpec {
  projectId: string;
  languages: string[];
  /**
   * `cached` mimics the snapshot texts Docentric stores for D365/ISV labels used by the
   * schema (@SYS…, @DocentricAX:…): no name, no section, usually one language only.
   */
  labels: Array<{ code: string; texts: Record<string, string>; cached?: boolean }>;
  dataSources: FixtureDataSource[];
  elements: FixtureElement[];
  /** Content controls with no element behind them, e.g. pasted from another template. */
  extraContentControls?: Array<{ tag: string; alias: string }>;
  /** Place these element ids a second time at the end of the body. */
  duplicatePlacements?: string[];
  reportId?: string;
  dspClass?: string;
}

const ENUM_TYPES = `
<xs:complexType name="Enum">
  <xs:attribute name="Value" type="xs:integer" use="optional"/>
  <xs:attribute name="Name" type="xs:string" use="optional"/>
  <xs:attribute name="Text" type="xs:string" use="optional"/>
</xs:complexType>`;

/** A small sales invoice: header + lines, company data, two languages. Validates clean. */
export function contosoInvoice(): FixtureSpec {
  return {
    projectId: '70000001',
    languages: ['en-us', 'it'],
    reportId: 'SalesInvoice.Report',
    dspClass: 'ConSalesInvoiceDSP',
    labels: [
      { code: '@DOC_LABEL:InvoiceTitle', texts: { 'en-us': 'Invoice', it: 'Fattura' } },
      { code: '@DOC_LABEL:Customer', texts: { 'en-us': 'Customer', it: 'Cliente' } },
      { code: '@DOC_LABEL:Total', texts: { 'en-us': 'Total', it: 'Totale' } },
      { code: '@SYS12836', texts: { 'en-us': 'Item number' }, cached: true },
      { code: '@DOC_LABEL:Discount', texts: { 'en-us': 'Discount', it: 'Sconto' } },
    ],
    dataSources: [
      {
        name: 'GeneralData',
        section: 2,
        schema: `${ENUM_TYPES}
<xs:element name="GeneralData"><xs:complexType><xs:choice minOccurs="0" maxOccurs="unbounded">
  <xs:element name="CurrentCompany"><xs:complexType>
    <xs:attribute name="Name" type="xs:string"/><xs:attribute name="VATNum" type="xs:string"/>
  </xs:complexType></xs:element>
  <xs:element name="Misc"><xs:complexType><xs:attribute name="LanguageId" type="xs:string"/></xs:complexType></xs:element>
</xs:choice></xs:complexType></xs:element>`,
        sample:
          '<GeneralData><CurrentCompany Name="Contoso Ltd" VATNum="XX000000"/><Misc LanguageId="en-us"/></GeneralData>',
      },
      {
        name: 'MainData',
        section: 1,
        schema: `${ENUM_TYPES}
<xs:element name="ReportData"><xs:complexType>
  <xs:choice minOccurs="0" maxOccurs="unbounded">
    <xs:element maxOccurs="unbounded" name="InvoiceHeader"><xs:complexType>
      <xs:choice minOccurs="0" maxOccurs="unbounded"><xs:element name="PrintLevel" type="Enum"/></xs:choice>
      <xs:attribute name="InvoiceId" type="xs:string"/>
      <xs:attribute name="InvoiceDate" type="xs:date"/>
      <xs:attribute name="CustName" type="xs:string"/>
      <xs:attribute name="CurrencyCode" type="xs:string"/>
      <xs:attribute name="InvoiceAmount" type="xs:decimal"/>
    </xs:complexType></xs:element>
    <xs:element maxOccurs="unbounded" name="InvoiceLines"><xs:complexType>
      <xs:attribute name="ItemId" type="xs:string"/>
      <xs:attribute name="Name" type="xs:string"/>
      <xs:attribute name="Qty" type="xs:decimal"/>
      <xs:attribute name="SalesPrice" type="xs:decimal"/>
      <xs:attribute name="LineAmount" type="xs:decimal"/>
      <xs:attribute name="DiscPercent" type="xs:decimal"/>
    </xs:complexType></xs:element>
  </xs:choice>
  <xs:attribute name="IsWithAmount" type="xs:boolean"/>
</xs:complexType></xs:element>`,
        sample:
          '<ReportData IsWithAmount="true"><InvoiceHeader InvoiceId="INV-0001" CustName="Fabrikam" CurrencyCode="EUR"/>' +
          '<InvoiceLines ItemId="A100" Name="Widget" Qty="2" SalesPrice="10" LineAmount="20" DiscPercent="0"/></ReportData>',
      },
      {
        name: 'Parameters',
        section: 3,
        schema: `${ENUM_TYPES}
<xs:element name="Parameters"><xs:complexType>
  <xs:attribute name="DocumentTitle" type="xs:string"/><xs:attribute name="IsProforma" type="xs:boolean"/>
</xs:complexType></xs:element>`,
        sample: '<Parameters DocumentTitle="Invoice" IsProforma="false"/>',
      },
    ],
    elements: [
      { id: '100', type: 'Field', source: 'Parameters', path: '@DocumentTitle', inHeader: true },
      { id: '101', type: 'Label', labelCode: '@DOC_LABEL:InvoiceTitle' },
      { id: '102', type: 'Field', source: 'MainData', path: 'InvoiceHeader/@InvoiceId' },
      {
        id: '103',
        type: 'Field',
        source: 'MainData',
        path: 'concat(label("@DOC_LABEL:Customer"), ": ", InvoiceHeader/@CustName)',
      },
      { id: '104', type: 'Field', source: 'GeneralData', path: 'upper-case(CurrentCompany/@Name)' },
      {
        id: '110',
        type: 'List',
        source: 'MainData',
        path: 'InvoiceLines',
        attrs: { RenderMode: '2' },
        children: [
          { id: '111', type: 'Field', source: '.', path: '@ItemId' },
          { id: '112', type: 'Field', source: '.', path: 'concat(@Name, $new-line, @ItemId)' },
          { id: '113', type: 'Field', source: '.', path: '@Qty * @SalesPrice', attrs: { FormatString: 'n2' } },
          {
            id: '114',
            type: 'If',
            source: '.',
            path: '@DiscPercent > 0',
            children: [
              { id: '115', type: 'Label', labelCode: '@DOC_LABEL:Discount' },
              { id: '116', type: 'Field', source: '.', path: 'round-number(@DiscPercent, 2)' },
            ],
          },
        ],
      },
      { id: '120', type: 'Var', name: 'tot', source: 'MainData', path: '0' },
      {
        id: '121',
        type: 'List',
        source: 'MainData',
        path: 'InvoiceLines',
        attrs: { RenderMode: '1' },
        children: [
          { id: '122', type: 'Var', name: 'tot', source: '.', path: "@LineAmount + var-element-value('tot')" },
        ],
      },
      {
        id: '130',
        type: 'If',
        source: 'MainData',
        path: '@IsWithAmount and InvoiceHeader/PrintLevel/@Value != 2',
        children: [
          { id: '131', type: 'Label', labelCode: '@DOC_LABEL:Total' },
          {
            id: '132',
            type: 'SummaryField',
            source: 'MainData',
            path: 'InvoiceLines',
            aggregatePath: '@Qty * @SalesPrice',
            attrs: { AggregateFunction: '1', FormatString: 'n2' },
          },
          { id: '133', type: 'Field', source: 'MainData', path: "round-number(var-element-value('tot'), 2)" },
        ],
      },
    ],
  };
}

/** Deep copy, so a test can mutate the spec without touching other tests. */
export function cloneSpec(spec: FixtureSpec): FixtureSpec {
  return structuredClone(spec);
}

export function findElement(spec: FixtureSpec, id: string): FixtureElement {
  const walk = (list: FixtureElement[]): FixtureElement | undefined => {
    for (const e of list) {
      if (e.id === id) return e;
      const hit = walk(e.children ?? []);
      if (hit) return hit;
    }
    return undefined;
  };
  const hit = walk(spec.elements);
  if (!hit) throw new Error(`fixture has no element ${id}`);
  return hit;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function flatten(list: FixtureElement[]): FixtureElement[] {
  return list.flatMap((e) => [e, ...flatten(e.children ?? [])]);
}

function projectDataXml(spec: FixtureSpec): string {
  const T = 'http://schemas.docentric.com/dynamics-ax/templates';
  const labels = spec.labels
    .map((l) => {
      const name = l.code.split(':').pop() ?? l.code;
      const texts = Object.entries(l.texts)
        .map(([lang, v]) => `<text d2p1:language="${esc(lang)}" d2p1:value="${esc(v)}"/>`)
        .join('');
      return l.cached
        ? `<label d2p1:code="${esc(l.code)}">${texts}</label>`
        : `<label d2p1:name="${esc(name)}" d2p1:code="${esc(l.code)}" d2p1:section="Custom">${texts}</label>`;
    })
    .join('');
  const languages = spec.languages.map((l) => `<language d2p1:value="${esc(l)}"/>`).join('');
  const dataSources = spec.dataSources
    .map(
      (ds) =>
        `<DataSource Name="${esc(ds.name)}" IsDefault="0" DataKind="2" ns1:dataSectionType="${ds.section}">` +
        `<Schema><XsdList><MainXsd xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:schema>${ds.schema}</xs:schema></MainXsd></XsdList>` +
        '<XmlNamespaceMappings/></Schema>' +
        `<PreviewData><SampleXml>${ds.sample ?? ''}</SampleXml></PreviewData></DataSource>`,
    )
    .join('');
  const elements = flatten(spec.elements)
    .map((e) => {
      const attrs = Object.entries(e.attrs ?? {})
        .map(([k, v]) => ` ${k}="${esc(v)}"`)
        .join('');
      const name = e.name ? ` Name="${esc(e.name)}"` : '';
      const childIds = [...(e.children ?? []).map((c) => c.id), ...(e.danglingChildIds ?? [])];
      const kids = childIds.length
        ? `<ChildElements>${childIds.map((id) => `<ElementRef Id="${id}"/>`).join('')}</ChildElements>`
        : '';
      const value = e.labelCode ? `<Value LabelCode="${esc(e.labelCode)}" IsCodeManuallyEntered="1"/>` : '';
      const binding =
        e.path !== undefined
          ? `<ValueBinding DataKind="2"${e.source ? ` Source="${esc(e.source)}"` : ''} Path="${esc(e.path)}" ResultDataType="System.String"/>`
          : '';
      const aggregate =
        e.aggregatePath !== undefined
          ? `<SortDescriptors/><AggregateBinding DataKind="2" Path="${esc(e.aggregatePath)}" ResultDataType="System.Double"/>`
          : '';
      return `<Element Type="${esc(e.type)}" Id="${e.id}"${name}${attrs}>${kids}${value}${binding}${aggregate}</Element>`;
    })
    .join('');

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' +
    '<ProjectData ProductId="3" ProductVersion="6.1.62.0" xmlns="http://Docentric.org/ProjectData.xml">' +
    '<UISettings xmlns=""><ElementsPane IsVisible="1"/></UISettings>' +
    `<dataSourcePackage d2p1:uri="fixture.ddsp" xmlns:d2p1="${T}" xmlns="${T}">` +
    `<ssrsReportInfo d2p1:reportId="${esc(spec.reportId ?? 'Fixture.Report')}" d2p1:reportGroupId="DEFAULT" ` +
    `d2p1:dataSourceProviderClassName="${esc(spec.dspClass ?? 'FixtureDSP')}" d2p1:dataSourceType="SSRS"/>` +
    `<languages d2p1:defaultLanguage="${esc(spec.languages[0] ?? 'en-us')}">${languages}</languages>` +
    `<labels>${labels}</labels></dataSourcePackage>` +
    `<DataSources xmlns="" xmlns:ns1="${T}">${dataSources}</DataSources>` +
    `<Elements xmlns="">${elements}</Elements>` +
    '</ProjectData>'
  );
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function sdt(tag: string, alias: string, text: string, inner = ''): string {
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${esc(alias)}"/><w:tag w:val="${esc(tag)}"/></w:sdtPr>` +
    `<w:sdtContent><w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>${inner}</w:sdtContent></w:sdt>`
  );
}

function renderElement(spec: FixtureSpec, e: FixtureElement): string {
  const inner = (e.children ?? []).map((c) => renderElement(spec, c)).join('');
  if (e.placed === false) return inner;
  const text = e.labelCode ?? (e.path ? `[${e.path}]` : e.type);
  return sdt(`${spec.projectId}-${e.id}`, e.alias ?? e.type, text, inner);
}

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function wordPart(root: 'document' | 'hdr', body: string): string {
  const open = root === 'document' ? `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` : `<w:hdr xmlns:w="${W}">`;
  // The section must reference the header part, or Word (and the Designer) never show it.
  const close =
    root === 'document'
      ? '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>'
      : '</w:hdr>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${open}${body || '<w:p/>'}${close}`;
}

export async function buildTemplate(spec: FixtureSpec): Promise<Buffer> {
  const zip = new JSZip();
  const all = flatten(spec.elements);
  const body =
    spec.elements
      .filter((e) => !e.inHeader)
      .map((e) => renderElement(spec, e))
      .join('') +
    (spec.extraContentControls ?? []).map((c) => sdt(c.tag, c.alias, c.alias)).join('') +
    (spec.duplicatePlacements ?? [])
      .map((id) => {
        const e = all.find((x) => x.id === id);
        return e ? sdt(`${spec.projectId}-${id}`, e.type, `${e.type} copy`) : '';
      })
      .join('');
  const header = spec.elements
    .filter((e) => e.inHeader)
    .map((e) => renderElement(spec, e))
    .join('');

  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
      '</Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
      '</Relationships>',
  );
  zip.file('word/document.xml', wordPart('document', body));
  zip.file('word/header1.xml', wordPart('hdr', header));
  zip.file('customXml/item1.xml', projectDataXml(spec));
  zip.file(
    'customXml/itemProps1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' +
      '<ds:datastoreItem ds:itemID="{00000000-0000-0000-0000-000000000001}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
      '<ds:schemaRefs><ds:schemaRef ds:uri="http://Docentric.org/ProjectData.xml"/></ds:schemaRefs></ds:datastoreItem>',
  );
  zip.file(
    'customXml/_rels/item1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>' +
      '</Relationships>',
  );
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
