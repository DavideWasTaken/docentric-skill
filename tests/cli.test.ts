import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { runCli, UsageError } from '../src/cli.js';
import { buildTemplate, cloneSpec, contosoInvoice, findElement } from './fixtures/builder.js';

let dir: string;
let output: string[];

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dcx-cli-'));
  await writeFile(join(dir, 'clean.docx'), await buildTemplate(contosoInvoice()));
  const broken = cloneSpec(contosoInvoice());
  findElement(broken, '102').path = 'InvoiceHeader/@Nope';
  await writeFile(join(dir, 'broken.docx'), await buildTemplate(broken));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});
beforeEach(() => {
  output = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => {
    output.push(a.join(' '));
  });
});
afterEach(() => {
  vi.restoreAllMocks();
});

describe('cli', () => {
  it('validates with an exit code CI can use', async () => {
    expect(await runCli(['validate', join(dir, 'clean.docx')])).toBe(0);
    expect(await runCli(['validate', join(dir, 'broken.docx'), '--min', 'error'])).toBe(1);
    expect(output.join('\n')).toContain('DCX007');
  });

  it('dumps a stable text view for git diff', async () => {
    await runCli(['dump', join(dir, 'clean.docx')]);
    const text = output.join('\n');
    expect(text).toContain('## Elements');
    expect(text).toContain('## Template labels');
  });

  it('creates from a YAML spec and previews to HTML', async () => {
    const spec = join(dir, 'spec.yaml');
    await writeFile(spec, 'dataSources: { from: clean.docx }\nbody:\n  - text: "No. {{InvoiceHeader/@InvoiceId}}"\n');
    expect(await runCli(['create', spec, join(dir, 'made.docx')])).toBe(0);
    expect(await runCli(['preview', join(dir, 'made.docx'), '--formats', 'html'])).toBe(0);
    expect(await readFile(join(dir, 'made.preview.html'), 'utf8')).toContain('No. INV-0001');
  });

  it('reads a spec or operations from stdin', async () => {
    const spec = ['dataSources: { from: clean.docx }', 'body:', '  - text: "{{InvoiceHeader/@CustName}}"'].join('\n');
    const prev = process.cwd();
    process.chdir(dir);
    try {
      expect(await runCli(['create', '-', join(dir, 'stdin.docx')], async () => spec)).toBe(0);
    } finally {
      process.chdir(prev);
    }
    const ops = JSON.stringify([{ op: 'add_language', language: 'de' }]);
    expect(
      await runCli(['edit', join(dir, 'stdin.docx'), '-', '--out', join(dir, 'stdin-de.docx')], async () => ops),
    ).toBe(0);
    expect(output.join('\n')).toMatch(/add_language: de added/);
  });

  it('edits from a JSON ops file', async () => {
    const ops = join(dir, 'ops.json');
    await writeFile(ops, JSON.stringify([{ op: 'set_label', code: 'X', texts: { 'en-us': 'x' } }]));
    expect(await runCli(['edit', join(dir, 'clean.docx'), ops, '--out', join(dir, 'edited.docx')])).toBe(0);
    expect(output.join('\n')).toContain('Saved');
  });

  it('prints help and rejects unknown commands', async () => {
    expect(await runCli(['--help'])).toBe(0);
    expect(output.join('\n')).toContain('convert-ssrs');
    await expect(runCli(['frobnicate'])).rejects.toBeInstanceOf(UsageError);
    await expect(runCli(['validate'])).rejects.toThrow(/usage/);
  });
});
