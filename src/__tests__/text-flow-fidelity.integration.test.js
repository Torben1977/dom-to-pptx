import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import puppeteer from 'puppeteer';
import JSZip from 'jszip';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { pathToFileURL } from 'node:url';
import { exportHtmlToPptx } from '../node-exporter.js';

// Compares where Office puts every word against where the browser put it. The
// fixture is OrgLith adapter output (see the note at its top), so the converter
// sees exactly what production hands it.

const runOfficeRoundtrip = process.env.DOM_TO_PPTX_OFFICE_ROUNDTRIP === '1';
const officeDescribe = runOfficeRoundtrip ? describe : describe.skip;

const FIXTURE = path.resolve('src/__tests__/fixtures/text-flow-fidelity-deck.html');
const CANVAS = { width: 1280, height: 720 };
const PX_TO_PT = 0.75;
const X_TOLERANCE_PT = 3;
const Y_TOLERANCE_PT = 4;
const SPACING_TOLERANCE_PT = 1;
const EMU_PER_PT = 12700;

// Defect kinds the converter shows today, per probe. A faithful probe has none.
// A converter fix removes its kind here; a regression adds one. Keep the kinds,
// not the measured numbers: numbers vary with fonts and browser builds.
const KNOWN_DEFECTS = {
  // The markers sit beside their text. What remains is the last item wrapping
  // onto a third line in Office because its text box got the width of its
  // content instead of the width available to it, which strands that item's
  // marker and pushes the list past its box.
  'marker-list': ['overflow', 'stranded'],
  'break-list': [],
  // The marker width comes out of the list's left inset, as it does out of the
  // left padding in the browser, so the text starts at the content edge.
  'native-bullets': [],
  // A negative text-indent travels as marL/indent with no marker. Faithful today.
  'hanging-indent': [],
  // The adapter materializes an inline `::before` as an inline span again, so it
  // joins the paragraph's run instead of getting a text box of its own.
  'inline-lead': [],
  'stacked-divs': [],
  // Row heights travel from the measured layout and the cell margins are the
  // CSS padding, nothing added. Faithful today.
  'table-cells': [],
  'table-below': [],
  // A span shifted with `position: relative` stays in the run on purpose: the
  // alternative is one box per fragment, which overlaps. Faithful today.
  'relative-offset': [],
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
let browserPages;
let officePages;
let slideXml;
let findingsByProbe;

const normalize = (word) => word.toLocaleLowerCase('de');

function groupLines(words) {
  const sorted = [...words].sort((a, b) => (a.top + a.bottom) / 2 - (b.top + b.bottom) / 2 || a.x - b.x);
  const lines = [];
  for (const word of sorted) {
    const center = (word.top + word.bottom) / 2;
    const line = lines.find((candidate) => Math.abs(candidate.center - center) < (word.bottom - word.top) / 2);
    if (line) line.words.push(word);
    else lines.push({ center, words: [word] });
  }
  for (const line of lines) line.words.sort((a, b) => a.x - b.x);
  return lines
    .map((line) => ({ words: line.words, top: Math.min(...line.words.map((word) => word.top)) }))
    .sort((a, b) => a.top - b.top);
}

async function measureBrowserPages(executablePath) {
  const browser = await puppeteer.launch({ executablePath, headless: true, args: ['--no-sandbox'] });
  try {
    const page = await browser.newPage();
    await page.setViewport(CANVAS);
    await page.goto(pathToFileURL(FIXTURE).href, { waitUntil: 'load' });
    await page.evaluate(() => document.fonts.ready);
    const pages = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.slide'), (slide) => {
        const origin = slide.getBoundingClientRect();
        const words = [];
        const walker = document.createTreeWalker(slide, NodeFilter.SHOW_TEXT);
        let node;
        while ((node = walker.nextNode())) {
          const probe = node.parentElement.closest('[data-probe]')?.dataset.probe ?? null;
          const pattern = /\S+/g;
          let match;
          while ((match = pattern.exec(node.textContent))) {
            const range = document.createRange();
            range.setStart(node, match.index);
            range.setEnd(node, match.index + match[0].length);
            const rect = Array.from(range.getClientRects()).find((candidate) => candidate.width > 0);
            if (!rect) continue;
            words.push({
              text: match[0],
              probe,
              x: rect.left - origin.left,
              xMax: rect.right - origin.left,
              top: rect.top - origin.top,
              bottom: rect.bottom - origin.top,
            });
          }
        }
        const probes = Array.from(slide.querySelectorAll('[data-probe]'), (element) => {
          const box = element.getBoundingClientRect();
          const style = getComputedStyle(element);
          const contentTop = box.top + parseFloat(style.borderTopWidth) + parseFloat(style.paddingTop);
          const contentBottom = box.bottom - parseFloat(style.borderBottomWidth) - parseFloat(style.paddingBottom);
          // Only in-flow block children open and close paragraphs with spacing.
          const blocks = Array.from(element.children).filter((child) => {
            const childStyle = getComputedStyle(child);
            return (
              childStyle.display !== 'none' &&
              !childStyle.display.startsWith('inline') &&
              !['absolute', 'fixed'].includes(childStyle.position) &&
              childStyle.float === 'none'
            );
          });
          const first = blocks[0]?.getBoundingClientRect();
          const last = blocks.at(-1)?.getBoundingClientRect();
          return {
            name: element.dataset.probe,
            top: box.top - origin.top,
            bottom: box.bottom - origin.top,
            left: box.left - origin.left,
            right: box.right - origin.left,
            // Space the first/last child leaves inside the content box; collapsed
            // margins lie outside the box and therefore count as zero here.
            leadingSpace: first ? first.top - contentTop : 0,
            trailingSpace: last ? contentBottom - last.bottom : 0,
          };
        });
        return { words, probes };
      })
    );
    const toPt = (value) => value * PX_TO_PT;
    return pages.map((page) => ({
      words: page.words.map((word) => ({
        ...word,
        x: toPt(word.x),
        xMax: toPt(word.xMax),
        top: toPt(word.top),
        bottom: toPt(word.bottom),
      })),
      probes: page.probes.map((probe) => ({
        ...probe,
        top: toPt(probe.top),
        bottom: toPt(probe.bottom),
        left: toPt(probe.left),
        right: toPt(probe.right),
        leadingSpace: toPt(probe.leadingSpace),
        trailingSpace: toPt(probe.trailingSpace),
      })),
    }));
  } finally {
    await browser.close();
  }
}

function parseOfficeWords(bboxHtml) {
  return bboxHtml
    .split('<page')
    .slice(1)
    .map((page) =>
      Array.from(
        page.matchAll(/<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)<\/word>/g),
        (match) => ({
          text: match[5].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'),
          x: Number(match[1]),
          top: Number(match[2]),
          xMax: Number(match[3]),
          bottom: Number(match[4]),
        })
      )
    );
}

/**
 * Pairs every browser word with the nearest Office word of the same text, once
 * each. Nearest-first keeps repeated words (markers, "der", "und") with their
 * own occurrence instead of whichever object mentions them first.
 */
function assignWords(browserWords, officeWords) {
  const pairs = [];
  browserWords.forEach((browserWord, browserIndex) => {
    officeWords.forEach((officeWord, officeIndex) => {
      if (normalize(browserWord.text) !== normalize(officeWord.text)) return;
      const distance = Math.hypot(officeWord.x - browserWord.x, officeWord.top - browserWord.top);
      pairs.push({ browserIndex, officeIndex, distance });
    });
  });
  pairs.sort((a, b) => a.distance - b.distance);
  const officeFor = new Map();
  const taken = new Set();
  for (const { browserIndex, officeIndex } of pairs) {
    if (officeFor.has(browserIndex) || taken.has(officeIndex)) continue;
    officeFor.set(browserIndex, officeIndex);
    taken.add(officeIndex);
  }
  return officeFor;
}

function textFrameParagraphs(xml, marker) {
  const shape = Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (match) => match[0]).find((candidate) =>
    Array.from(candidate.matchAll(/<a:t>([^<]*)<\/a:t>/g), (run) => run[1]).some((text) => text.includes(marker))
  );
  if (!shape) return null;
  return Array.from(shape.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g), (match) => {
    const spacing = (tag) => {
      const value = match[1].match(new RegExp(`<a:${tag}><a:spcPts val="(\\d+)"`));
      return value ? Number(value[1]) / 100 : 0;
    };
    return { before: spacing('spcBef'), after: spacing('spcAft') };
  });
}

/**
 * Differences a reader would see between the browser and Office, as
 * `{ probe, kind, detail }`. Wrapping at a different word is not one by itself.
 */
function analyzePage(browserPage, officeWords, xml) {
  const findings = [];
  const add = (probe, kind, detail) => {
    if (probe) findings.push({ probe, kind, detail });
  };
  const officeFor = assignWords(browserPage.words, officeWords);
  const probeOf = new Map();
  officeFor.forEach((officeIndex, browserIndex) => probeOf.set(officeIndex, browserPage.words[browserIndex].probe));

  const browserLineOf = new Map();
  for (const probe of browserPage.probes) {
    const indexed = browserPage.words
      .map((word, index) => ({ ...word, index }))
      .filter((word) => word.probe === probe.name);
    groupLines(indexed).forEach((line, lineIndex) => {
      line.words.forEach((word, position) =>
        browserLineOf.set(word.index, { lineIndex, first: position === 0, last: position === line.words.length - 1 })
      );
    });
  }

  browserPage.words.forEach((word, index) => {
    if (word.probe in RASTERIZED) return;
    if (!officeFor.has(index)) add(word.probe, 'missing-word', `"${word.text}" is missing or broken in Office`);
  });

  for (const line of groupLines(officeWords.map((word, index) => ({ ...word, index })))) {
    for (let position = 0; position < line.words.length - 1; position++) {
      const [left, right] = [line.words[position], line.words[position + 1]];
      if (left.xMax <= right.x + 1) continue;
      const detail = `"${left.text}" overlaps "${right.text}"`;
      const probes = new Set([probeOf.get(left.index), probeOf.get(right.index)]);
      probes.forEach((probe) => add(probe, 'overlap', detail));
    }
  }

  const officeByProbe = new Map();
  for (const probe of browserPage.probes.filter((candidate) => !(candidate.name in RASTERIZED))) {
    const pairs = browserPage.words
      .map((word, index) => ({ word, index }))
      .filter(({ word, index }) => word.probe === probe.name && officeFor.has(index))
      .map(({ word, index }) => ({
        ...officeWords[officeFor.get(index)],
        browser: word,
        line: browserLineOf.get(index),
      }));
    officeByProbe.set(probe.name, pairs);

    const bottom = Math.max(...pairs.map((word) => word.bottom));
    if (bottom > probe.bottom + Y_TOLERANCE_PT) {
      add(probe.name, 'overflow', `text ends ${(bottom - probe.bottom).toFixed(1)} pt below its browser box`);
    }

    const officeLines = groupLines(pairs);
    for (const line of officeLines) {
      const first = line.words[0];
      if (first.line.first && Math.abs(first.x - first.browser.x) > X_TOLERANCE_PT) {
        add(
          probe.name,
          'line-start',
          `line starting "${first.text}" at x=${first.x.toFixed(1)} pt, browser x=${first.browser.x.toFixed(1)} pt`
        );
      }
      // A word alone on its Office line although the browser continued its line
      // after it has been torn from its text, e.g. a marker.
      if (line.words.length === 1 && !first.line.last) {
        add(probe.name, 'stranded', `"${first.text}" stands alone on its line`);
      }
    }

    const offsets = [];
    for (const browserLine of new Set(pairs.map((word) => word.line.lineIndex))) {
      const members = pairs.filter((word) => word.line.lineIndex === browserLine);
      const officeLine = officeLines.find(
        (line) => line.words.length === members.length && members.every((word) => line.words.includes(word))
      );
      if (officeLine) offsets.push({ text: members[0].text, offset: officeLine.top - members[0].browser.top });
    }
    for (const { text, offset } of offsets.slice(1)) {
      const drift = offset - offsets[0].offset;
      if (Math.abs(drift) > Y_TOLERANCE_PT) {
        add(probe.name, 'drift', `line starting "${text}" drifts ${drift.toFixed(1)} pt vertically`);
      }
    }

    const marker = [...pairs.map((word) => word.browser.text)].sort((a, b) => b.length - a.length)[0];
    const paragraphs = marker ? textFrameParagraphs(xml, marker) : null;
    if (paragraphs?.length) {
      const leading = paragraphs[0].before;
      const trailing = paragraphs[paragraphs.length - 1].after;
      if (Math.abs(leading - probe.leadingSpace) > SPACING_TOLERANCE_PT) {
        add(
          probe.name,
          'edge-spacing',
          `first paragraph spacing ${leading} pt, browser ${probe.leadingSpace.toFixed(1)} pt`
        );
      }
      if (Math.abs(trailing - probe.trailingSpace) > SPACING_TOLERANCE_PT) {
        add(
          probe.name,
          'edge-spacing',
          `last paragraph spacing ${trailing} pt, browser ${probe.trailingSpace.toFixed(1)} pt`
        );
      }
    }
  }

  // Objects stacked apart in the browser must not run into each other in Office.
  const stackable = browserPage.probes.filter((candidate) => !(candidate.name in RASTERIZED));
  for (const upper of stackable) {
    for (const lower of stackable) {
      const stacked = upper.bottom <= lower.top && upper.left < lower.right && lower.left < upper.right;
      const upperWords = officeByProbe.get(upper.name) || [];
      const lowerWords = officeByProbe.get(lower.name) || [];
      if (!stacked || !upperWords.length || !lowerWords.length) continue;
      const upperBottom = Math.max(...upperWords.map((word) => word.bottom));
      const lowerTop = Math.min(...lowerWords.map((word) => word.top));
      if (upperBottom > lowerTop + 1) {
        const detail = `"${upper.name}" runs ${(upperBottom - lowerTop).toFixed(1)} pt into "${lower.name}"`;
        add(upper.name, 'collision', detail);
        add(lower.name, 'collision', detail);
      }
    }
  }

  return findings;
}

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
    browserPages = await measureBrowserPages(executablePath);

    outputDir = mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-text-flow-fidelity-'));
    const pptxPath = path.join(outputDir, 'text-flow-fidelity.pptx');
    const pdfPath = path.join(outputDir, 'text-flow-fidelity.pdf');
    const bboxPath = path.join(outputDir, 'text-flow-fidelity.html');
    const buffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      // The policy the controlled deck path is meant to run with: an object the
      // converter cannot map becomes a picture instead of failing the export.
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false, boundaryPolicy: 'rasterize' },
    });
    writeFileSync(pptxPath, buffer);
    const zip = await JSZip.loadAsync(buffer);
    slideXml = await Promise.all(
      browserPages.map((_, index) => zip.file(`ppt/slides/slide${index + 1}.xml`).async('string'))
    );
    execFileSync(
      'soffice',
      [
        `-env:UserInstallation=${pathToFileURL(path.join(outputDir, 'libreoffice-profile')).href}`,
        '--headless',
        '--convert-to',
        'pdf',
        '--outdir',
        outputDir,
        pptxPath,
      ],
      { stdio: 'pipe' }
    );
    execFileSync('pdftotext', ['-bbox', pdfPath, bboxPath], { stdio: 'pipe' });
    officePages = parseOfficeWords(readFileSync(bboxPath, 'utf8'));

    findingsByProbe = new Map([...Object.keys(KNOWN_DEFECTS), ...Object.keys(RASTERIZED)].map((probe) => [probe, []]));
    browserPages.forEach((page, index) => {
      for (const finding of analyzePage(page, officePages[index], slideXml[index])) {
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

  it.each(Object.entries(RASTERIZED))('hands %s over as a picture covering its box (%s)', (probe) => {
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
  });
});
