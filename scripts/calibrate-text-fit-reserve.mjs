import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../src/node-exporter.js';

const EMU_PER_POINT = 12_700;
const CASE_IDS = ['R00', 'R01', 'R02', 'R03', 'R04', 'R05', 'R06'];
const DEFAULT_RESERVES = [0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06];

function requireExecutable(command, versionArgs = ['--version']) {
  try {
    execFileSync(command, versionArgs, { stdio: 'ignore' });
  } catch (error) {
    throw new Error(`Text-fit calibration requires '${command}' on PATH.`, { cause: error });
  }
}

function parseReserves() {
  const configured = process.env.DOM_TO_PPTX_TEXT_FIT_RESERVES;
  if (!configured) return DEFAULT_RESERVES;

  const reserves = configured.split(',').map((value) => Number(value.trim()));
  if (reserves.length === 0 || reserves.some((value) => !Number.isFinite(value) || value < 0 || value > 0.25)) {
    throw new Error('DOM_TO_PPTX_TEXT_FIT_RESERVES must contain comma-separated ratios between 0 and 0.25.');
  }
  return reserves;
}

function decodeXml(value) {
  return value
    .replaceAll('&amp;', '&')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'");
}

function attribute(xml, name) {
  const match = xml.match(new RegExp(`\\b${name}="([^"]+)"`));
  return match ? Number(match[1]) : 0;
}

function parsePdfLines(xml) {
  const page = xml.match(/<page\b[^>]*>([\s\S]*?)<\/page>/)?.[1] || '';
  return Array.from(page.matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/g), ([, attrs, body]) => ({
    text: Array.from(body.matchAll(/<word\b[^>]*>([\s\S]*?)<\/word>/g), (match) => decodeXml(match[1])).join(' '),
    yMin: attribute(attrs, 'yMin'),
    yMax: attribute(attrs, 'yMax'),
  }));
}

async function analyzePptx(buffer, bboxXml) {
  const zip = await JSZip.loadAsync(buffer);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  const shapes = Array.from(slideXml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (match) => match[0]);
  const lines = parsePdfLines(bboxXml);

  return CASE_IDS.map((id) => {
    const shape = shapes.find((candidate) => candidate.includes(`<a:t>${id}</a:t>`));
    if (!shape) throw new Error(`Calibration shape '${id}' is missing from slide XML.`);

    const transform = shape.match(/<a:xfrm[^>]*>([\s\S]*?)<\/a:xfrm>/)?.[1] || '';
    const offset = transform.match(/<a:off\b([^>]*)\/>/)?.[1] || '';
    const extent = transform.match(/<a:ext\b([^>]*)\/>/)?.[1] || '';
    const bodyPr = shape.match(/<a:bodyPr\b([^>]*)>/)?.[1] || '';
    const line = lines.find((candidate) => candidate.text === id);
    if (!line) throw new Error(`LibreOffice-rendered line '${id}' is missing from PDF text output.`);

    const top = attribute(offset, 'y') / EMU_PER_POINT;
    const height = attribute(extent, 'cy') / EMU_PER_POINT;
    const contentTop = top + attribute(bodyPr, 'tIns') / EMU_PER_POINT;
    const contentBottom = top + height - attribute(bodyPr, 'bIns') / EMU_PER_POINT;
    const tolerance = 1;

    return {
      id,
      autofit: shape.includes('<a:normAutofit'),
      contentTop,
      contentBottom,
      renderedTop: line.yMin,
      renderedBottom: line.yMax,
      renderedHeight: line.yMax - line.yMin,
      fits: line.yMin >= contentTop - tolerance && line.yMax <= contentBottom + tolerance,
    };
  });
}

function printResults(results) {
  const header = ['reserve', ...CASE_IDS.map((id) => `${id} fit/auto`)];
  const rows = results.map(({ reserve, cases }) => [
    `${Math.round(reserve * 100)}%`,
    ...cases.map((entry) => `${entry.fits ? 'yes' : 'NO'}/${entry.autofit ? 'yes' : 'no'}`),
  ]);
  const widths = header.map((value, index) => Math.max(value.length, ...rows.map((row) => String(row[index]).length)));
  const format = (row) => row.map((value, index) => String(value).padEnd(widths[index])).join('  ');

  console.log(format(header));
  console.log(format(widths.map((width) => '-'.repeat(width))));
  for (const row of rows) console.log(format(row));

  const baseline = results.find(({ reserve }) => reserve === 0);
  if (!baseline) return;
  const baselineFailures = baseline.cases.filter((entry) => !entry.fits);
  if (baselineFailures.length === 0) {
    console.log('\nLibreOffice result: the calibration cases do not demonstrate a need for a reserve above 0%.');
    return;
  }

  const firstPassing = results.find(({ cases }) => cases.every((entry) => entry.fits));
  if (firstPassing) {
    console.log(
      `\nLibreOffice result: the smallest tested reserve with all cases visually inside their text boxes is ${Math.round(firstPassing.reserve * 100)}%.`
    );
  } else {
    console.log(
      '\nLibreOffice result: none of the tested reserves keeps all calibration cases inside their text boxes.'
    );
  }
}

requireExecutable('soffice');
requireExecutable('pdftotext', ['-v']);

const fixture = path.resolve('src/__tests__/fixtures/text-fit-reserve-calibration.html');
const configuredOutput = process.env.DOM_TO_PPTX_TEXT_FIT_OUTPUT_DIR;
const outputDir = configuredOutput
  ? path.resolve(configuredOutput)
  : mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-text-fit-calibration-'));
mkdirSync(outputDir, { recursive: true });

const results = [];
try {
  for (const reserve of parseReserves()) {
    const label = `reserve-${String(Math.round(reserve * 100)).padStart(2, '0')}`;
    const variantDir = path.join(outputDir, label);
    mkdirSync(variantDir, { recursive: true });
    const pptxPath = path.join(variantDir, `${label}.pptx`);
    const pdfPath = path.join(variantDir, `${label}.pdf`);
    const bboxPath = path.join(variantDir, `${label}.xml`);
    const profilePath = path.join(variantDir, 'libreoffice-profile');

    const buffer = await exportHtmlToPptx(fixture, {
      selector: '.slide',
      pptxOptions: {
        width: 13.333333,
        height: 7.5,
        autoEmbedFonts: false,
        _textFitReserveRatio: reserve,
      },
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
        variantDir,
        pptxPath,
      ],
      { stdio: 'pipe' }
    );
    if (!existsSync(pdfPath)) throw new Error(`LibreOffice did not produce ${pdfPath}.`);
    execFileSync('pdftotext', ['-bbox-layout', pdfPath, bboxPath], { stdio: 'pipe' });
    const cases = await analyzePptx(buffer, readFileSync(bboxPath, 'utf8'));
    results.push({ reserve, cases });
  }

  printResults(results);
  writeFileSync(path.join(outputDir, 'results.json'), `${JSON.stringify(results, null, 2)}\n`);
  console.log(`\nCalibration artifacts: ${outputDir}`);
} finally {
  if (!configuredOutput) rmSync(outputDir, { recursive: true, force: true });
}
