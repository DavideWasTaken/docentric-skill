import { randomUUID } from 'node:crypto';
import { WordPackage } from '../docx/package.js';
import { esc, solidPng } from './ooxml.js';

/**
 * Assembles a complete Word package around generated body/header/footer XML and
 * a ProjectData part. With a base document, its styles, theme, fonts and
 * numbering are reused so generated templates match the corporate look.
 */

export interface AssembleInput {
  body: string;
  header?: string;
  footer?: string;
  projectDataXml: string;
  page: { width: number; height: number; landscape: boolean; top: number; right: number; bottom: number; left: number };
  font: { name: string; size: number };
  title: string;
  /** Image placeholder token to replace with the part's relationship id. */
  imageToken?: string;
  base?: WordPackage;
  /** Custom document properties (e.g. provisional data source marker). */
  customProperties?: Record<string, string>;
}

const CT = {
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  fontTable: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  customXmlProps: 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  custom: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
};

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const DOC_NAMESPACES =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'mc:Ignorable="w14 w15"';

export async function assemble(input: AssembleInput): Promise<WordPackage> {
  const pkg = WordPackage.empty();
  const xmlDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const overrides: Array<[string, string]> = [
    ['word/document.xml', CT.document],
    ['word/styles.xml', CT.styles],
    ['word/settings.xml', CT.settings],
    ['customXml/itemProps1.xml', CT.customXmlProps],
    ['docProps/core.xml', CT.core],
    ['docProps/app.xml', CT.app],
  ];
  const docRels: Array<[string, string]> = [
    [`${REL}/styles`, 'styles.xml'],
    [`${REL}/settings`, 'settings.xml'],
    [`${REL}/customXml`, '../customXml/item1.xml'],
  ];

  // Styles: from the base document, or a clean default set.
  const baseStyles = input.base ? await input.base.text('word/styles.xml') : null;
  pkg.setText('word/styles.xml', baseStyles ?? stylesXml(input.font));
  if (input.base) {
    for (const [part, ct, rel] of [
      ['word/theme/theme1.xml', CT.theme, 'theme'],
      ['word/fontTable.xml', CT.fontTable, 'fontTable'],
      ['word/numbering.xml', CT.numbering, 'numbering'],
    ] as const) {
      const xml = await input.base.text(part);
      if (!xml) continue;
      pkg.setText(part, xml);
      overrides.push([part, ct]);
      docRels.push([`${REL}/${rel}`, part.replace('word/', '')]);
    }
  }
  pkg.setText(
    'word/settings.xml',
    `${xmlDecl}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      '<w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      '</w:settings>',
  );

  // Header / footer parts.
  let headerRef = '';
  let footerRef = '';
  const usesImage = (xml?: string) => !!(xml && input.imageToken && xml.includes(input.imageToken));
  const anyImage = usesImage(input.body) || usesImage(input.header) || usesImage(input.footer);
  if (anyImage) {
    pkg.setBytes('word/media/placeholder.png', solidPng(16, 8, [217, 217, 217]));
  }
  const partWithImages = (xml: string, relsTarget: string): string => {
    if (!usesImage(xml)) return xml;
    pkg.setText(
      relsTarget,
      `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImg1" Type="${REL}/image" Target="media/placeholder.png"/></Relationships>`,
    );
    return xml.split(input.imageToken as string).join('rIdImg1');
  };
  if (input.header) {
    const xml = partWithImages(input.header, 'word/_rels/header1.xml.rels');
    pkg.setText('word/header1.xml', `${xmlDecl}<w:hdr ${DOC_NAMESPACES}>${xml || '<w:p/>'}</w:hdr>`);
    overrides.push(['word/header1.xml', CT.header]);
    docRels.push([`${REL}/header`, 'header1.xml']);
    headerRef = `<w:headerReference w:type="default" r:id="rId${docRels.length}"/>`;
  }
  if (input.footer) {
    const xml = partWithImages(input.footer, 'word/_rels/footer1.xml.rels');
    pkg.setText('word/footer1.xml', `${xmlDecl}<w:ftr ${DOC_NAMESPACES}>${xml || '<w:p/>'}</w:ftr>`);
    overrides.push(['word/footer1.xml', CT.footer]);
    docRels.push([`${REL}/footer`, 'footer1.xml']);
    footerRef = `<w:footerReference w:type="default" r:id="rId${docRels.length}"/>`;
  }

  // Body.
  let body = input.body;
  if (usesImage(body)) {
    docRels.push([`${REL}/image`, 'media/placeholder.png']);
    body = body.split(input.imageToken as string).join(`rId${docRels.length}`);
  }
  const p = input.page;
  const sectPr =
    `<w:sectPr>${headerRef}${footerRef}` +
    `<w:pgSz w:w="${p.width}" w:h="${p.height}"${p.landscape ? ' w:orient="landscape"' : ''}/>` +
    `<w:pgMar w:top="${p.top}" w:right="${p.right}" w:bottom="${p.bottom}" w:left="${p.left}" w:header="567" w:footer="567" w:gutter="0"/>` +
    '<w:cols w:space="708"/></w:sectPr>';
  pkg.setText(
    'word/document.xml',
    `${xmlDecl}<w:document ${DOC_NAMESPACES}><w:body>${body}${sectPr}</w:body></w:document>`,
  );
  pkg.setText(
    'word/_rels/document.xml.rels',
    `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      docRels
        .map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${target}"/>`)
        .join('') +
      '</Relationships>',
  );

  // Docentric ProjectData.
  pkg.setText('customXml/item1.xml', input.projectDataXml);
  pkg.setText(
    'customXml/itemProps1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' +
      `<ds:datastoreItem ds:itemID="{${randomUUID().toUpperCase()}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
      '<ds:schemaRefs><ds:schemaRef ds:uri="http://Docentric.org/ProjectData.xml"/><ds:schemaRef ds:uri=""/>' +
      '<ds:schemaRef ds:uri="http://schemas.docentric.com/dynamics-ax/templates"/><ds:schemaRef ds:uri="http://www.w3.org/2001/XMLSchema"/>' +
      '<ds:schemaRef ds:uri="http://schemas.docentric.com/xml-schema-extensions"/></ds:schemaRefs></ds:datastoreItem>',
  );
  pkg.setText(
    'customXml/_rels/item1.xml.rels',
    `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
  );

  // Package properties.
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  pkg.setText(
    'docProps/core.xml',
    `${xmlDecl}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${esc(input.title)}</dc:title><dc:creator>docentric-skill</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  pkg.setText(
    'docProps/app.xml',
    `${xmlDecl}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
      '<Application>docentric-skill</Application></Properties>',
  );
  const pkgRels: Array<[string, string]> = [
    [`${REL}/officeDocument`, 'word/document.xml'],
    ['http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'],
    [`${REL}/extended-properties`, 'docProps/app.xml'],
  ];
  if (input.customProperties && Object.keys(input.customProperties).length) {
    pkg.setText('docProps/custom.xml', customPropertiesXml(input.customProperties));
    overrides.push(['docProps/custom.xml', CT.custom]);
    pkgRels.push([`${REL}/custom-properties`, 'docProps/custom.xml']);
  }
  pkg.setText(
    '_rels/.rels',
    `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      pkgRels
        .map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${target}"/>`)
        .join('') +
      '</Relationships>',
  );

  pkg.setText(
    '[Content_Types].xml',
    `${xmlDecl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      (anyImage ? '<Default Extension="png" ContentType="image/png"/>' : '') +
      overrides.map(([part, ct]) => `<Override PartName="/${part}" ContentType="${ct}"/>`).join('') +
      '</Types>',
  );
  return pkg;
}

export function customPropertiesXml(props: Record<string, string>): string {
  const items = Object.entries(props)
    .map(
      ([name, value], i) =>
        `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${esc(name)}"><vt:lpwstr>${esc(value)}</vt:lpwstr></property>`,
    )
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${items}</Properties>`
  );
}

export function stylesXml(font: { name: string; size: number }): string {
  const sz = Math.round(font.size * 2);
  const heading = (n: number, pt: number, before: number) =>
    `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/>` +
    `<w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="${before}" w:after="80"/>` +
    `<w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="${pt * 2}"/><w:szCs w:val="${pt * 2}"/></w:rPr></w:style>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    `<w:rFonts w:ascii="${esc(font.name)}" w:hAnsi="${esc(font.name)}" w:eastAsia="${esc(font.name)}" w:cs="${esc(font.name)}"/>` +
    `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>` +
    '<w:pPrDefault><w:pPr><w:spacing w:after="40" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    heading(1, 16, 240) +
    heading(2, 13, 200) +
    heading(3, 11, 160) +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/>' +
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '</w:styles>'
  );
}
