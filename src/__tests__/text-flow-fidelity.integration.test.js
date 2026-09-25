import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import puppeteer from 'puppeteer';
import JSZip from 'jszip';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { exportHtmlToPptx } from '../node-exporter.js';
import {
  convertToPdf,
  EMU_PER_PT,
  X_TOLERANCE_PT,
  Y_TOLERANCE_PT,
  analyzePage,
  measureBrowserPages,
  parseOfficeWords,
} from './helpers/office-fidelity.js';

// Compares where Office puts every word against where the browser put it. The
// fixture is OrgLith adapter output (see the note at its top), so the converter
// sees exactly what production hands it.

const runOfficeRoundtrip = process.env.DOM_TO_PPTX_OFFICE_ROUNDTRIP === '1';
const officeDescribe = runOfficeRoundtrip ? describe : describe.skip;

const FIXTURE = path.resolve('src/__tests__/fixtures/text-flow-fidelity-deck.html');
const CANVAS = { width: 1280, height: 720 };

// Defect kinds the converter shows today, per probe. A faithful probe has none.
// A converter fix removes its kind here; a regression adds one. Keep the kinds,
// not the measured numbers: numbers vary with fonts and browser builds.
const KNOWN_DEFECTS = {
  // The markers sit beside their text, and the wrapped items get the width the
  // browser wrapped them in. This is the pattern that broke in production.
  'marker-list': [],
  'break-list': [],
  // The marker width comes out of the list's left inset, as it does out of the
  // left padding in the browser, so the text starts at the content edge.
  'native-bullets': [],
  // A negative text-indent travels as marL/indent with no marker. Faithful today.
  'hanging-indent': [],
  // The adapter materializes an inline `::before` as an inline span again, so it
  // joins the paragraph's run instead of getting a text box of its own. The
  // paragraph shares one text box with the heading, as wide as the slide, and
  // loses its own 560 px: Office pulls up words the browser wrapped.
  'inline-lead': ['rewrap'],
  'stacked-divs': [],
  // Row heights travel from the measured layout and the cell margins are the
  // CSS padding, nothing added. Faithful today.
  'table-cells': [],
  'table-below': [],
  // A span shifted with `position: relative` stays in the run on purpose: the
  // alternative is one box per fragment, which overlaps. The paragraph loses its
  // 560 px in the slide-wide text box, as in `inline-lead`.
  'relative-offset': ['rewrap'],
};

// Probes the converter is expected to hand over as a picture instead of text,
// because PowerPoint has no way to hold them as text at all. This is a result,
// not a defect — but it costs editable text, so it has to be declared here and
// proven: no words in the text layer, and a picture covering the object's box.
// Without this distinction the oracle cannot tell a fix from a cop-out, since a
// rasterized object and two lost words both read as `missing-word`.
const RASTERIZED = {
  // A float is the one out-of-flow box that moves the rest of the text: the
  // lines beside it are shortened, the lines below it are not. One PowerPoint
  // rectangle cannot be both.
  'float-marker': 'float-in-text-flow',
  // A native PowerPoint table cell holds text and nothing else, so the cell's
  // absolutely positioned marker has nowhere to go. The cell cannot be replaced
  // on its own, so the table is the smallest object that can be.
  'table-marker': 'table-cell-needs-shape',
};

let outputDir;
let boundaryFindings;
let browserPages;
let officePages;
let slideXml;
let findingsByProbe;
let fixtureSlideIds;

const kindsOf = (probe) => [...new Set(findingsByProbe.get(probe).map((finding) => finding.kind))].sort();

officeDescribe('Office text flow fidelity against the browser layout', () => {
  beforeAll(async () => {
    // The exporter launches Puppeteer's bundled browser; measure with the same build.
    const executablePath = await puppeteer.executablePath();
    if (!existsSync(executablePath)) {
      throw new Error(
        `Puppeteer's browser is missing at ${executablePath}; run 'npx puppeteer browsers install chrome'.`
      );
    }
    ({ pages: browserPages, slideIds: fixtureSlideIds } = await measureBrowserPages(executablePath, FIXTURE, CANVAS));

    boundaryFindings = [];
    outputDir = mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-text-flow-fidelity-'));
    const pptxPath = path.join(outputDir, 'text-flow-fidelity.pptx');
    const pdfPath = path.join(outputDir, 'text-flow-fidelity.pdf');
    const bboxPath = path.join(outputDir, 'text-flow-fidelity.html');
    const buffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      // The policy the controlled deck path is meant to run with: an object the
      // converter cannot map becomes a picture instead of failing the export.
      pptxOptions: {
        width: 13.333333,
        height: 7.5,
        autoEmbedFonts: false,
        boundaryPolicy: 'rasterize',
        onBoundaryFindings: (findings) => boundaryFindings.push(...findings),
      },
    });
    writeFileSync(pptxPath, buffer);
    const zip = await JSZip.loadAsync(buffer);
    slideXml = await Promise.all(
      browserPages.map((_, index) => zip.file(`ppt/slides/slide${index + 1}.xml`).async('string'))
    );
    convertToPdf(pptxPath, outputDir);
    execFileSync('pdftotext', ['-bbox', pdfPath, bboxPath], { stdio: 'pipe' });
    officePages = parseOfficeWords(readFileSync(bboxPath, 'utf8'));

    findingsByProbe = new Map([...Object.keys(KNOWN_DEFECTS), ...Object.keys(RASTERIZED)].map((probe) => [probe, []]));
    browserPages.forEach((page, index) => {
      for (const finding of analyzePage(page, officePages[index], slideXml[index], RASTERIZED)) {
        findingsByProbe.get(finding.probe)?.push(finding);
      }
    });
  }, 120_000);

  afterAll(() => {
    if (outputDir) rmSync(outputDir, { recursive: true, force: true });
  });

  it('measures exactly the declared probes, each in the browser and in Office', () => {
    const measured = browserPages.flatMap((page) => page.probes.map((probe) => probe.name));
    const declared = [...Object.keys(KNOWN_DEFECTS), ...Object.keys(RASTERIZED)];
    expect(new Set(declared).size, 'a probe is either measured as text or declared as a picture').toBe(declared.length);
    expect(measured.sort()).toEqual(declared.sort());
    expect(officePages).toHaveLength(browserPages.length);
    for (const page of browserPages) {
      for (const probe of page.probes) {
        expect(
          page.words.some((word) => word.probe === probe.name),
          `browser words for ${probe.name}`
        ).toBe(true);
      }
    }
  });

  it.each(Object.entries(KNOWN_DEFECTS))('shows exactly the known defect kinds for %s', (probe, expected) => {
    expect(kindsOf(probe), JSON.stringify(findingsByProbe.get(probe), null, 2)).toEqual(expected);
  });

  it.each(Object.entries(RASTERIZED))('hands %s over as a picture covering its box (%s)', (probe, expectedType) => {
    const pageIndex = browserPages.findIndex((page) => page.probes.some((candidate) => candidate.name === probe));
    const box = browserPages[pageIndex].probes.find((candidate) => candidate.name === probe);
    const expectedWords = browserPages[pageIndex].words.filter((word) => word.probe === probe).map((word) => word.text);
    expect(expectedWords.length, `${probe} carries text in the browser`).toBeGreaterThan(0);

    // Nothing may be left in the text layer where the object sits, or it is both
    // a picture and a text box and the reader sees the object twice. Asked of
    // the box rather than of the words, because the same words legitimately
    // occur elsewhere on the slide — the heading names the probe.
    const strays = officePages[pageIndex].filter((word) => {
      const centerX = (word.x + word.xMax) / 2;
      const centerY = (word.top + word.bottom) / 2;
      return (
        centerX > box.left - X_TOLERANCE_PT &&
        centerX < box.right + X_TOLERANCE_PT &&
        centerY > box.top - Y_TOLERANCE_PT &&
        centerY < box.bottom + Y_TOLERANCE_PT
      );
    });
    expect(
      strays.map((word) => word.text),
      `${probe} must not also appear as text`
    ).toEqual([]);

    const pictures = Array.from(slideXml[pageIndex].matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g), (match) => {
      const offset = match[0].match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
      const extent = match[0].match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
      if (!offset || !extent) return null;
      return {
        left: Number(offset[1]) / EMU_PER_PT,
        top: Number(offset[2]) / EMU_PER_PT,
        right: (Number(offset[1]) + Number(extent[1])) / EMU_PER_PT,
        bottom: (Number(offset[2]) + Number(extent[2])) / EMU_PER_PT,
      };
    }).filter(Boolean);

    const covering = pictures.find(
      (picture) =>
        picture.left <= box.left + X_TOLERANCE_PT &&
        picture.top <= box.top + Y_TOLERANCE_PT &&
        picture.right >= box.right - X_TOLERANCE_PT &&
        picture.bottom >= box.bottom - Y_TOLERANCE_PT
    );
    expect(
      covering,
      `no picture covers ${probe} at ${JSON.stringify(box)}; pictures: ${JSON.stringify(pictures)}`
    ).toBeDefined();

    // And for the declared reason. A picture in the right place proves only that
    // something replaced the object; if it was replaced for another reason, the
    // detector this probe exists for is not the one that fired.
    const slideId = fixtureSlideIds[pageIndex];
    const reported = boundaryFindings.filter((finding) => finding.slideId === slideId).map((finding) => finding.type);
    expect(reported, `${probe} on slide ${slideId} was replaced for another reason`).toEqual([expectedType]);
  });
});
