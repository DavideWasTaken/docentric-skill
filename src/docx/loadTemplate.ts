import { extractContentControls, isContentPart } from '../model/contentControls.js';
import { isProjectData, parseProjectData } from '../model/projectData.js';
import type { ContentControl, DocentricTemplate } from '../model/types.js';
import { parseXml } from '../xml/dom.js';
import { readTemplateBytes, TemplateError, WordPackage } from './package.js';

export { TemplateError };

/** Loads a Docentric template from disk. The file is only read, never written. */
export async function loadTemplateFile(filePath: string): Promise<DocentricTemplate> {
  return loadTemplate(await readTemplateBytes(filePath));
}

export async function loadTemplate(data: Buffer | Uint8Array | WordPackage): Promise<DocentricTemplate> {
  const pkg = data instanceof WordPackage ? data : await WordPackage.open(data);

  const partName = await findProjectDataPart(pkg);
  if (!partName) {
    throw new TemplateError(
      'No Docentric ProjectData part found (customXml/item*.xml). This is a Word file but not a Docentric template.',
    );
  }
  const projectData = parseProjectData(parseXml((await pkg.text(partName)) as string, partName), partName);

  const contentControls: ContentControl[] = [];
  for (const part of contentParts(pkg)) {
    const xml = await pkg.text(part);
    if (xml) contentControls.push(...extractContentControls(parseXml(xml, part), part));
  }

  return { ...projectData, contentControls, provisionalDataSources: await readProvisional(pkg) };
}

/** Written by the create command when a data source schema was declared rather than imported from a DDSP. */
async function readProvisional(pkg: WordPackage): Promise<string[]> {
  const xml = await pkg.text('docProps/custom.xml');
  const m = xml
    ? /name="DocentricSkill\.ProvisionalDataSources"[^>]*>\s*<vt:lpwstr>([^<]*)<\/vt:lpwstr>/.exec(xml)
    : null;
  return m?.[1] ? m[1].split(',').filter(Boolean) : [];
}

export async function findProjectDataPart(pkg: WordPackage): Promise<string | null> {
  for (const name of pkg.parts().filter((n) => /^customXml\/item\d+\.xml$/.test(n))) {
    const xml = await pkg.text(name);
    if (!xml?.includes('ProjectData')) continue;
    if (isProjectData(parseXml(xml, name))) return name;
  }
  return null;
}

/** document.xml first, then headers, footers, notes — the order a reader meets them. */
export function contentParts(pkg: WordPackage): string[] {
  const rank = (p: string) => (p.includes('document') ? 0 : p.includes('header') ? 1 : p.includes('footer') ? 2 : 3);
  return pkg
    .parts()
    .filter(isContentPart)
    .sort((a, b) => rank(a) - rank(b) || a.localeCompare(b, undefined, { numeric: true }));
}
