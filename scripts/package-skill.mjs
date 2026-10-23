// Packs skills/docentric into dist/docentric-skill.zip for claude.ai (Settings → Capabilities → Skills → Upload).
// The zip holds one top-level folder, `docentric/`, with SKILL.md at its root. Test cases (evals/) are left out.
import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { join, relative, sep } from 'node:path';
import JSZip from 'jszip';

const root = 'skills/docentric';
const zip = new JSZip();
const skip = (rel) => rel.startsWith('evals');

async function add(dir) {
  for (const name of await readdir(dir)) {
    const full = join(dir, name);
    const rel = relative(root, full).split(sep).join('/');
    if (skip(rel)) continue;
    if ((await stat(full)).isDirectory()) await add(full);
    else zip.file(`docentric/${rel}`, await readFile(full), { date: new Date('2026-01-01T00:00:00Z') });
  }
}

await add(root);
await mkdir('dist', { recursive: true });
const out = 'dist/docentric-skill.zip';
await writeFile(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`${out} — ${Object.keys(zip.files).filter((f) => !zip.files[f].dir).join(', ')}`);
