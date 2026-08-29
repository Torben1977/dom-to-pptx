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
let shadowPixelCounts;

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

function lineContaining(pageNumber, expectedText) {
  const line = pageLines(pageNumber).find(({ text }) => text.includes(expectedText));
  expect(line, `one Office-rendered line containing '${expectedText}' on page ${pageNumber}`).toBeDefined();
  return line;
}

function readPpmPixels(filePath) {
  const buffer = readFileSync(filePath);
  let offset = 0;
  const whitespace = (byte) => byte === 9 || byte === 10 || byte === 13 || byte === 32;
  const readToken = () => {
    while (offset < buffer.length) {
      if (whitespace(buffer[offset])) {
        offset++;
        continue;
      }
      if (buffer[offset] === 35) {
        while (offset < buffer.length && buffer[offset] !== 10) offset++;
        continue;
      }
      break;
    }
    const start = offset;
    while (offset < buffer.length && !whitespace(buffer[offset])) offset++;
    return buffer.subarray(start, offset).toString('ascii');
  };

  expect(readToken()).toBe('P6');
  const width = Number(readToken());
  const height = Number(readToken());
  expect(Number(readToken())).toBe(255);
  if (buffer[offset] === 13 && buffer[offset + 1] === 10) offset += 2;
  else if (whitespace(buffer[offset])) offset++;
  const pixels = buffer.subarray(offset);
  expect(pixels.length).toBe(width * height * 3);
  return pixels;
}

officeDescribe('LibreOffice visual round trip', () => {
  beforeAll(async () => {
    requireExecutable('soffice');
    requireExecutable('pdftotext', ['-v']);
    requireExecutable('pdftoppm', ['-v']);

    outputDir = mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-office-roundtrip-'));
    const fixture = path.resolve('src/__tests__/fixtures/orglith-text-flow-deck.html');
    const pptxPath = path.join(outputDir, 'orglith-text-flow.pptx');
    const pdfPath = path.join(outputDir, 'orglith-text-flow.pdf');
    const bboxPath = path.join(outputDir, 'orglith-text-flow.xml');
    const shadowRasterPrefix = path.join(outputDir, 'shadow-svg-render');
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
    execFileSync('pdftoppm', ['-f', '6', '-l', '6', '-singlefile', '-r', '96', pdfPath, shadowRasterPrefix], {
      stdio: 'pipe',
    });

    documentNode = new DOMParser().parseFromString(readFileSync(bboxPath, 'utf8'), 'text/xml');
    const pixels = readPpmPixels(`${shadowRasterPrefix}.ppm`);
    let teal = 0;
    let darkRed = 0;
    for (let index = 0; index < pixels.length; index += 3) {
      const red = pixels[index];
      const green = pixels[index + 1];
      const blue = pixels[index + 2];
      if (red < 80 && green > 85 && blue > 70 && blue < 165) teal++;
      if (red > 120 && green < 80 && blue < 80) darkRed++;
    }
    shadowPixelCounts = { teal, darkRed };
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

  it('renders the Shadow-DOM SVG and its decorated host through LibreOffice', () => {
    expect(shadowPixelCounts.teal).toBeGreaterThan(1_000);
    expect(shadowPixelCounts.darkRed).toBeGreaterThan(500);
  });

  it('keeps browser-fit mixed text at readable Office-rendered sizes', () => {
    const pillTitle = lineWith(7, 'Gemeinsame');
    const pillSubtitle = lineWith(7, 'statt Aktivität');
    const timelineTitle = lineWith(7, 'Verbindlicher Steuerungsrahmen');
    const timelineBody = lineStartingWith(7, 'Ergebnisse, Kapazität');

    expect(pillTitle.yMax - pillTitle.yMin).toBeGreaterThan(12);
    expect(pillSubtitle.yMax - pillSubtitle.yMin).toBeGreaterThan(8);
    expect(timelineTitle.yMax - timelineTitle.yMin).toBeGreaterThan(12);
    expect(timelineBody.yMax - timelineBody.yMin).toBeGreaterThan(8);
    expect(timelineBody.yMin).toBeGreaterThan(timelineTitle.yMax);
  });

  it('keeps metric values and currency units on one readable Office-rendered line', () => {
    const primaryMetric = lineWith(8, '2,6 Mio. €');
    const costMetric = lineWith(8, '4,1 Mio. €');
    const support = lineWith(8, 'höher als im Business Case angenommen');
    const capacity = lineWith(8, '5–10 %');
    const duration = lineWith(8, '12 Monate');

    expect(primaryMetric.yMax - primaryMetric.yMin).toBeGreaterThan(30);
    expect(costMetric.yMax - costMetric.yMin).toBeGreaterThan(18);
    // Glyph bounding boxes can overlap slightly even when their baselines and
    // painted pixels do not. Guard the actual vertical progression instead of
    // assuming disjoint font ascender/descender boxes.
    expect(support.yMin).toBeGreaterThan(costMetric.yMin + 20);
    expect(capacity.yMax - capacity.yMin).toBeGreaterThan(10);
    expect(duration.yMax - duration.yMin).toBeGreaterThan(7);
  });

  it('keeps a list footer below the final bullet after Office lays out the card', () => {
    const finalBullet = lineContaining(9, 'Eine Zusage der beteiligten Personalvertretungen liegt noch nicht vor.');
    const note = lineWith(9, 'Diese Punkte sind offene Unsicherheiten, keine stillschweigenden Annahmen.');

    expect(note.yMin).toBeGreaterThan(finalBullet.yMax);
  });

  it('keeps compact axis labels on one Office-rendered line', () => {
    const first = lineWith(10, 'Wirksame, aber teure Einzellösung');
    const second = lineWith(10, 'Vertiefung · gesamter Wertstrom');

    expect(first.yMax - first.yMin).toBeGreaterThan(7);
    expect(second.yMax - second.yMin).toBeGreaterThan(8);
  });
});
