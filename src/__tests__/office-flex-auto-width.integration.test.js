import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { pathToFileURL } from 'node:url';
import { exportHtmlToPptx } from '../node-exporter.js';

const runOfficeRoundtrip = process.env.DOM_TO_PPTX_OFFICE_ROUNDTRIP === '1';
const officeDescribe = runOfficeRoundtrip ? describe : describe.skip;
const AUTO_LABELS = ['Entscheidung jetzt', 'Wechselwirkung', 'Ausgangswerte', 'Umsetzung', 'Review'];

let outputDir;
let documentNode;

function requireExecutable(command, versionArgs = ['--version']) {
  try {
    execFileSync(command, versionArgs, { stdio: 'ignore' });
  } catch (error) {
    throw new Error(`The Office round-trip contract requires '${command}' on PATH.`, { cause: error });
  }
}

function pageLines(pageNumber) {
  const page = documentNode.getElementsByTagName('page')[pageNumber - 1];
  expect(page, `rendered PDF page ${pageNumber}`).toBeDefined();
  return Array.from(page.getElementsByTagName('line'), (line) =>
    Array.from(line.getElementsByTagName('word'), (word) => word.textContent).join(' ')
  );
}

officeDescribe('LibreOffice auto-width flex round trip', () => {
  beforeAll(async () => {
    requireExecutable('soffice');
    requireExecutable('pdftotext', ['-v']);

    outputDir = mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-flex-auto-width-'));
    const fixture = path.resolve('src/__tests__/fixtures/flex-auto-width-regression.html');
    const pptxPath = path.join(outputDir, 'flex-auto-width.pptx');
    const pdfPath = path.join(outputDir, 'flex-auto-width.pdf');
    const bboxPath = path.join(outputDir, 'flex-auto-width.xml');
    const profilePath = path.join(outputDir, 'libreoffice-profile');

    const buffer = await exportHtmlToPptx(fixture, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    writeFileSync(pptxPath, buffer);
    execFileSync(
      'soffice',
      [
        `-env:UserInstallation=${pathToFileURL(profilePath).href}`,
        '--headless',
        '--convert-to',
        'pdf',
        '--outdir',
        outputDir,
        pptxPath,
      ],
      { stdio: 'pipe' }
    );
    execFileSync('pdftotext', ['-bbox-layout', pdfPath, bboxPath], { stdio: 'pipe' });
    documentNode = new DOMParser().parseFromString(readFileSync(bboxPath, 'utf8'), 'text/xml');
  }, 60_000);

  afterAll(() => {
    if (outputDir) rmSync(outputDir, { recursive: true, force: true });
  });

  it('keeps every painted auto-width flex label on one rendered line', () => {
    const lines = pageLines(1);
    for (const label of AUTO_LABELS) expect(lines, label).toContain(label);
  });

  it('preserves deliberate line breaks and normal word wrapping', () => {
    const lines = pageLines(3);
    expect(lines).toContain('Erste Zeile');
    expect(lines).toContain('Zweite Zeile');
    expect(lines).toContain('Dieser Text soll weiterhin');
    expect(lines).toContain('regulär zwischen Wörtern');
    expect(lines).toContain('umbrechen können.');
    expect(lines).toContain('Normaler Absatztext bleibt in seiner');
    expect(lines).toContain('begrenzten Textbox und verhält sich');
    expect(lines).toContain('unverändert.');
  });
});
