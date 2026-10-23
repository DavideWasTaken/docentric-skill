import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';

/**
 * External converters for the preview. All optional; each is looked up on the
 * machine and skipped when absent:
 *   LibreOffice (any OS)     .docx → .pdf, best fidelity after Word
 *   Microsoft Word (Windows) .docx → .pdf through COM automation
 *   Edge / Chrome / Chromium the HTML preview → .png screenshot, or → .pdf
 */

export type PdfEngine = 'libreoffice' | 'word' | 'browser';

const TIMEOUT_MS = 120_000;

function run(cmd: string, args: string[], timeout = TIMEOUT_MS): Promise<{ ok: boolean; output: string }> {
  return new Promise((resolve) => {
    execFile(cmd, args, { timeout, windowsHide: true, maxBuffer: 8 * 1024 * 1024 }, (err, stdout, stderr) => {
      resolve({ ok: !err, output: `${stdout ?? ''}${stderr ?? ''}${err ? `\n${err.message}` : ''}`.trim() });
    });
  });
}

function firstExisting(candidates: Array<string | undefined>): string | null {
  for (const c of candidates) if (c && existsSync(c)) return c;
  return null;
}

function onPath(names: string[]): string | null {
  const dirs = (process.env.PATH ?? '').split(process.platform === 'win32' ? ';' : ':');
  const exts = process.platform === 'win32' ? ['.exe', '.cmd', ''] : [''];
  for (const d of dirs)
    for (const n of names) for (const e of exts) if (d && existsSync(join(d, n + e))) return join(d, n + e);
  return null;
}

export function findLibreOffice(): string | null {
  const pf = process.env.ProgramFiles ?? 'C:\\Program Files';
  const pf86 = process.env['ProgramFiles(x86)'] ?? 'C:\\Program Files (x86)';
  return (
    firstExisting([
      process.env.DOCENTRIC_SOFFICE,
      join(pf, 'LibreOffice', 'program', 'soffice.exe'),
      join(pf86, 'LibreOffice', 'program', 'soffice.exe'),
      '/Applications/LibreOffice.app/Contents/MacOS/soffice',
      '/usr/bin/soffice',
      '/usr/bin/libreoffice',
      '/snap/bin/libreoffice',
    ]) ?? onPath(['soffice', 'libreoffice'])
  );
}

export function findBrowser(): string | null {
  const pf = process.env.ProgramFiles ?? 'C:\\Program Files';
  const pf86 = process.env['ProgramFiles(x86)'] ?? 'C:\\Program Files (x86)';
  const local = process.env.LOCALAPPDATA ?? '';
  return (
    firstExisting([
      process.env.DOCENTRIC_BROWSER,
      join(pf86, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
      join(pf, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
      join(pf, 'Google', 'Chrome', 'Application', 'chrome.exe'),
      join(pf86, 'Google', 'Chrome', 'Application', 'chrome.exe'),
      local && join(local, 'Google', 'Chrome', 'Application', 'chrome.exe'),
      '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    ]) ?? onPath(['google-chrome', 'chromium', 'chromium-browser', 'microsoft-edge'])
  );
}

export async function wordAvailable(): Promise<boolean> {
  if (process.platform !== 'win32') return false;
  const r = await run(
    'powershell.exe',
    [
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      "[bool](Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\\Word.Application\\CLSID' -ErrorAction SilentlyContinue)",
    ],
    20_000,
  );
  return r.ok && r.output.toLowerCase().includes('true');
}

export async function libreOfficeToPdf(soffice: string, docx: string, pdf: string): Promise<string | null> {
  const work = await mkdtemp(join(tmpdir(), 'dcx-lo-'));
  try {
    // A private profile avoids clashing with a LibreOffice instance the user has open.
    const profile = pathToFileURL(join(work, 'profile')).href;
    const r = await run(soffice, [
      `-env:UserInstallation=${profile}`,
      '--headless',
      '--convert-to',
      'pdf',
      '--outdir',
      work,
      docx,
    ]);
    const produced = join(work, basename(docx).replace(/\.[^.]+$/, '.pdf'));
    if (!existsSync(produced)) return `LibreOffice did not produce a PDF: ${r.output.slice(0, 300)}`;
    await copyFile(produced, pdf);
    return null;
  } finally {
    await rm(work, { recursive: true, force: true });
  }
}

/**
 * Word COM export. Runs in its own Word instance (`/Automation`), never the
 * user's. If Word stalls on a hidden dialog (sensitivity labels, recovery), the
 * timeout kills only automation instances started after this call.
 */
export async function wordToPdf(docx: string, pdf: string): Promise<string | null> {
  const ps = (s: string) => s.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$start = Get-Date
$job = Start-Job -ScriptBlock {
  param($src, $dst)
  $word = New-Object -ComObject Word.Application
  try {
    $word.Visible = $false; $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($src, $false, $true, $false)
    $doc.ExportAsFixedFormat($dst, 17)
    $doc.Close(0)
    'ok'
  } finally { $word.Quit() }
} -ArgumentList '${ps(docx)}', '${ps(pdf)}'
if (Wait-Job $job -Timeout 90) { Receive-Job $job } else {
  Stop-Job $job
  Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
    Where-Object { $_.CommandLine -like '*/Automation*' -and $_.CreationDate -ge $start } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
  'timeout'
}
Remove-Job $job -Force`;
  const r = await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], 150_000);
  if (existsSync(pdf)) return null;
  return r.output.includes('timeout')
    ? 'Word did not answer within 90 s (a hidden dialog, e.g. a sensitivity label prompt?).'
    : `Word export failed: ${r.output.slice(0, 300)}`;
}

export async function browserToPdf(browser: string, html: string, pdf: string): Promise<string | null> {
  const r = await run(browser, [
    '--headless=new',
    '--disable-gpu',
    '--no-pdf-header-footer',
    `--print-to-pdf=${pdf}`,
    pathToFileURL(html).href,
  ]);
  return existsSync(pdf) ? null : `Browser PDF export failed: ${r.output.slice(0, 300)}`;
}

/** Real height in pixels of an HTML page at the preview width, measured by the browser itself. */
export async function browserMeasureHeight(browser: string, html: string, width = 900): Promise<number | null> {
  const probe = `${html}.measure.html`;
  const script =
    "<script>document.documentElement.setAttribute('data-height', String(document.documentElement.scrollHeight))</script></body>";
  await writeFile(probe, (await readFile(html, 'utf8')).replace('</body>', script), 'utf8');
  try {
    const r = await run(browser, [
      '--headless=new',
      '--disable-gpu',
      `--window-size=${width},1250`,
      '--dump-dom',
      pathToFileURL(probe).href,
    ]);
    const m = /data-height="(\d+)"/.exec(r.output);
    return m ? Number(m[1]) : null;
  } finally {
    await rm(probe, { force: true });
  }
}

/** `height`: pixels to capture. */
export async function browserScreenshot(
  browser: string,
  html: string,
  png: string,
  height = 1250,
  width = 900,
): Promise<string | null> {
  const r = await run(browser, [
    '--headless=new',
    '--disable-gpu',
    '--hide-scrollbars',
    `--window-size=${width},${height}`,
    `--screenshot=${png}`,
    pathToFileURL(html).href,
  ]);
  return existsSync(png) ? null : `Browser screenshot failed: ${r.output.slice(0, 300)}`;
}

export function outputBase(templatePath: string, outDir?: string): string {
  const dir = outDir ?? dirname(templatePath);
  return join(dir, basename(templatePath).replace(/\.(docx|dotx|docm|dotm)$/i, ''));
}
