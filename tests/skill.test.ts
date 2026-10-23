import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { parse } from 'yaml';
import { VERSION } from '../src/version.js';

const SKILL_DIR = 'skills/docentric';
const BUNDLE = join(SKILL_DIR, 'scripts', 'docentric.mjs');
const skillMd = readFileSync(join(SKILL_DIR, 'SKILL.md'), 'utf8');
const frontmatter = parse(/^---\n([\s\S]*?)\n---/.exec(skillMd.replace(/\r\n/g, '\n'))?.[1] ?? '') as Record<
  string,
  string
>;

describe('SKILL.md', () => {
  it('has the frontmatter Claude needs to load and trigger it', () => {
    expect(frontmatter.name).toBe('docentric');
    expect(frontmatter.name).toMatch(/^[a-z0-9-]{1,64}$/);
    expect(frontmatter.description?.length).toBeGreaterThan(200);
    // Claude truncates longer descriptions; the trigger phrases must fit.
    expect(frontmatter.description?.length).toBeLessThanOrEqual(1024);
  });

  it('stays short: details belong in references/', () => {
    expect(skillMd.split('\n').length).toBeLessThan(200);
  });

  it('points only at files that exist', () => {
    for (const ref of skillMd.matchAll(/`((?:references|scripts)\/[\w.-]+)`/g)) {
      expect(existsSync(join(SKILL_DIR, ref[1] as string)), ref[1]).toBe(true);
    }
  });

  it('documents every edit operation the tool accepts', () => {
    const ops = readFileSync(join(SKILL_DIR, 'references', 'edit-operations.md'), 'utf8');
    for (const op of [
      'insert',
      'add_column',
      'set_binding',
      'wrap_if',
      'remove',
      'set_label',
      'add_language',
      'resize_column',
      'convert_placeholders',
    ]) {
      expect(ops).toContain(`\`${op}\``);
    }
  });
});

describe('bundled script', () => {
  const run = (...args: string[]) => execFileSync(process.execPath, [BUNDLE, ...args], { encoding: 'utf8' });

  it('is built from the current sources (run `npm run build`)', () => {
    expect(run('--version').trim()).toBe(VERSION);
  });

  it('runs a real command without node_modules', () => {
    expect(run('validate', 'examples/SalesInvoice.docx')).toContain('0 error(s)');
  });

  it('agrees with the plugin manifests on the version', () => {
    const plugin = JSON.parse(readFileSync('.claude-plugin/plugin.json', 'utf8'));
    const market = JSON.parse(readFileSync('.claude-plugin/marketplace.json', 'utf8'));
    expect(plugin.version).toBe(VERSION);
    expect(market.plugins[0].version).toBe(VERSION);
  });
});
