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

let outputDir;
let documentNode;

function requireExecutable(command, versionArgs = ['--version']) {
  try {
    execFileSync(command, versionArgs, { stdio: 'ignore' });
  } catch (error) {
    throw new Error(
      `The Office round-trip contract requires '${command}' on PATH. ` +
        `Install LibreOffice and Poppler before running npm run test:office-roundtrip.`,
      { cause: error }
    );
  }
}

function pageLines(pageNumber) {
  const page = documentNode.getElementsByTagName('page')[pageNumber - 1];
  expect(page, `rendered PDF page ${pageNumber}`).toBeDefined();
  return Array.from(page.getElementsByTagName('line'), (line) => ({
    text: Array.from(line.getElementsByTagName('word'), (word) => word.textContent).join(' '),
    yMin: Number(line.getAttribute('yMin')),
    yMax: Number(line.getAttribute('yMax')),
  }));
}

function lineWith(pageNumber, expectedText) {
  const line = pageLines(pageNumber).find(({ text }) => text === expectedText);
  expect(line, `one Office-rendered line '${expectedText}' on page ${pageNumber}`).toBeDefined();
  return line;
}

function lineStartingWith(pageNumber, expectedPrefix) {
  const line = pageLines(pageNumber).find(({ text }) => text.startsWith(expectedPrefix));
  expect(line, `one Office-rendered line starting with '${expectedPrefix}' on page ${pageNumber}`).toBeDefined();
  return line;
}

officeDescribe('LibreOffice visual text-flow round trip', () => {
  beforeAll(async () => {
    requireExecutable('soffice');
    requireExecutable('pdftotext', ['-v']);

    outputDir = mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-office-roundtrip-'));
    const fixture = path.resolve('src/__tests__/fixtures/orglith-text-flow-deck.html');
    const pptxPath = path.join(outputDir, 'orglith-text-flow.pptx');
    const pdfPath = path.join(outputDir, 'orglith-text-flow.pdf');
    const bboxPath = path.join(outputDir, 'orglith-text-flow.xml');
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

  it('keeps representative card headings on one rendered line', () => {
    lineWith(1, 'Gemeinsam messen');
    lineWith(1, 'Gemeinsam befähigen');
    lineWith(1, 'Nach 6 Monaten prüfen');
    lineWith(2, 'Weiter wie bisher');
    lineWith(2, 'Verbindlicher Rahmen');
  });

  it('keeps card body text below its heading after Office lays out the PPTX', () => {
    const heading = lineWith(2, 'Weiter wie bisher');
    const body = lineWith(2, 'Lokale Optimierung,');
    expect(body.yMin).toBeGreaterThan(heading.yMax);

    const secondHeading = lineWith(2, 'Verbindlicher Rahmen');
    const secondBody = lineStartingWith(2, 'Gemeinsame Ergebnisse,');
    expect(secondBody.yMin).toBeGreaterThan(secondHeading.yMax);
  });

  it('keeps open-point labels separate from their explanatory paragraphs', () => {
    const heading = lineWith(3, 'Einsparungen');
    const body = lineStartingWith(3, 'Es gibt noch keine belastbare Schätzung');
    expect(body.yMin).toBeGreaterThan(heading.yMax);

    const legalHeading = lineWith(3, 'Recht');
    const legalBody = lineStartingWith(3, 'Eine abschließende rechtliche Bewertung');
    expect(legalBody.yMin).toBeGreaterThan(legalHeading.yMax);
  });

  it('keeps metric values, units, and labels in their intended vertical order', () => {
    const value = lineWith(4, '1,8');
    const unit = lineWith(4, 'Mio.');
    const label = lineWith(4, 'Vorgänge pro Jahr');
    expect(unit.yMin).toBeGreaterThanOrEqual(value.yMax);
    expect(label.yMin).toBeGreaterThan(unit.yMax);

    const periodStart = lineWith(4, '2027–');
    const periodEnd = lineWith(4, '2029');
    const periodLabel = lineStartingWith(4, 'Zeitraum des');
    expect(periodEnd.yMin).toBeGreaterThanOrEqual(periodStart.yMax);
    expect(periodLabel.yMin).toBeGreaterThan(periodEnd.yMax);
  });
});
