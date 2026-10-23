import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, extname, resolve } from 'node:path';
import { type Document, DOMParser, type Element, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';

/**
 * Thin OPC (Office Open XML package) layer: read parts as text or DOM, write them
 * back, add parts with their content type and relationship. Everything that
 * changes a .docx goes through here so the package stays consistent.
 */

export class TemplateError extends Error {}

const MAX_BYTES = 50 * 1024 * 1024;
const WORD_EXTENSIONS = new Set(['.docx', '.dotx', '.docm', '.dotm']);

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  projectData: 'http://Docentric.org/ProjectData.xml',
  docTemplates: 'http://schemas.docentric.com/dynamics-ax/templates',
} as const;

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  customXml: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  customXmlProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
} as const;

export class WordPackage {
  private constructor(readonly zip: JSZip) {}

  static async open(data: Buffer | Uint8Array): Promise<WordPackage> {
    const zip = await JSZip.loadAsync(data).catch(() => {
      throw new TemplateError('The file is not a valid Word (OOXML) package.');
    });
    return new WordPackage(zip);
  }

  static async openFile(filePath: string): Promise<WordPackage> {
    return WordPackage.open(await readTemplateBytes(filePath));
  }

  static empty(): WordPackage {
    return new WordPackage(new JSZip());
  }

  parts(): string[] {
    return Object.keys(this.zip.files).filter((n) => !this.zip.files[n]?.dir);
  }

  has(part: string): boolean {
    return !!this.zip.file(part);
  }

  async text(part: string): Promise<string | null> {
    return (await this.zip.file(part)?.async('string')) ?? null;
  }

  async bytes(part: string): Promise<Uint8Array | null> {
    return (await this.zip.file(part)?.async('uint8array')) ?? null;
  }

  async dom(part: string): Promise<Document | null> {
    const xml = await this.text(part);
    return xml === null ? null : parseDocument(xml, part);
  }

  setText(part: string, xml: string): void {
    this.zip.file(part, xml);
  }

  setBytes(part: string, data: Uint8Array): void {
    this.zip.file(part, data);
  }

  setDom(part: string, doc: Document): void {
    this.zip.file(part, serialize(doc));
  }

  remove(part: string): void {
    this.zip.remove(part);
  }

  /** Registers a content-type override (or a default for an extension). */
  async setContentType(part: string, contentType: string): Promise<void> {
    const doc = (await this.dom('[Content_Types].xml')) ?? parseDocument(`<Types xmlns="${NS.ct}"/>`, 'ct');
    const root = doc.documentElement as Element;
    const name = `/${part}`;
    for (const o of Array.from(root.getElementsByTagName('Override'))) {
      if (o.getAttribute('PartName') === name) {
        o.setAttribute('ContentType', contentType);
        this.setDom('[Content_Types].xml', doc);
        return;
      }
    }
    const o = doc.createElementNS(NS.ct, 'Override');
    o.setAttribute('PartName', name);
    o.setAttribute('ContentType', contentType);
    root.appendChild(o);
    this.setDom('[Content_Types].xml', doc);
  }

  async ensureDefaultContentType(extension: string, contentType: string): Promise<void> {
    const doc = (await this.dom('[Content_Types].xml')) ?? parseDocument(`<Types xmlns="${NS.ct}"/>`, 'ct');
    const root = doc.documentElement as Element;
    for (const d of Array.from(root.getElementsByTagName('Default'))) {
      if (d.getAttribute('Extension')?.toLowerCase() === extension.toLowerCase()) return;
    }
    const d = doc.createElementNS(NS.ct, 'Default');
    d.setAttribute('Extension', extension);
    d.setAttribute('ContentType', contentType);
    root.insertBefore(d, root.firstChild);
    this.setDom('[Content_Types].xml', doc);
  }

  /** Adds a relationship from `sourcePart` and returns its id (reuses an identical one). */
  async addRelationship(sourcePart: string, type: string, target: string): Promise<string> {
    const relsPart = relsPartFor(sourcePart);
    const doc = (await this.dom(relsPart)) ?? parseDocument(`<Relationships xmlns="${NS.rel}"/>`, relsPart);
    const root = doc.documentElement as Element;
    const used = new Set<string>();
    for (const r of Array.from(root.getElementsByTagName('Relationship'))) {
      const id = r.getAttribute('Id') ?? '';
      used.add(id);
      if (r.getAttribute('Type') === type && r.getAttribute('Target') === target) return id;
    }
    let n = used.size + 1;
    while (used.has(`rId${n}`)) n++;
    const r = doc.createElementNS(NS.rel, 'Relationship');
    r.setAttribute('Id', `rId${n}`);
    r.setAttribute('Type', type);
    r.setAttribute('Target', target);
    root.appendChild(r);
    this.setDom(relsPart, doc);
    return `rId${n}`;
  }

  async relationships(sourcePart: string): Promise<Array<{ id: string; type: string; target: string }>> {
    const doc = await this.dom(relsPartFor(sourcePart));
    if (!doc) return [];
    return Array.from((doc.documentElement as Element).getElementsByTagName('Relationship')).map((r) => ({
      id: r.getAttribute('Id') ?? '',
      type: r.getAttribute('Type') ?? '',
      target: r.getAttribute('Target') ?? '',
    }));
  }

  async toBuffer(): Promise<Buffer> {
    return this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }

  async saveAs(filePath: string): Promise<string> {
    const full = resolve(filePath);
    assertAllowedPath(full);
    await mkdir(dirname(full), { recursive: true });
    await writeFile(full, await this.toBuffer());
    return full;
  }
}

export function relsPartFor(part: string): string {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

export function parseDocument(xml: string, partName: string): Document {
  const doc = new DOMParser({
    onError: (level, message) => {
      if (level !== 'warning') throw new TemplateError(`Malformed XML in ${partName}: ${message}`);
    },
  }).parseFromString(xml, 'text/xml');
  if (!doc.documentElement) throw new TemplateError(`Empty XML part: ${partName}`);
  return doc;
}

export function serialize(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc);
  return xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;
}

export async function readTemplateBytes(filePath: string): Promise<Buffer> {
  const full = resolve(filePath);
  assertAllowedPath(full);
  if (!WORD_EXTENSIONS.has(extname(full).toLowerCase())) {
    throw new TemplateError(`Not a Word package (expected .docx/.dotx/.docm/.dotm): ${full}`);
  }
  const info = await stat(full).catch(() => null);
  if (!info?.isFile()) throw new TemplateError(`File not found: ${full}`);
  if (info.size > MAX_BYTES) throw new TemplateError(`File is larger than ${MAX_BYTES} bytes: ${full}`);
  return readFile(full);
}

/**
 * Optional sandbox: DOCENTRIC_ALLOWED_ROOTS is a list of folders separated by
 * the platform path delimiter (';' on Windows). When set, only files under them
 * can be read or written.
 */
export function assertAllowedPath(full: string): void {
  const raw = process.env.DOCENTRIC_ALLOWED_ROOTS;
  if (!raw) return;
  const win = process.platform === 'win32';
  const sep = win ? '\\' : '/';
  const norm = (p: string) => (win ? resolve(p).toLowerCase() : resolve(p));
  const target = norm(full);
  const ok = raw
    .split(win ? ';' : ':')
    .filter(Boolean)
    .map(norm)
    .some((root) => target === root || target.startsWith(root.endsWith(sep) ? root : root + sep));
  if (!ok) throw new TemplateError(`Path is outside DOCENTRIC_ALLOWED_ROOTS: ${full}`);
}
