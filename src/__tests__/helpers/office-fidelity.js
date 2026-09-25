import puppeteer from 'puppeteer';
import { execFileSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { pathToFileURL } from 'node:url';

// Measures where every word and every probe box lies in the browser and where
// Office puts the same words, and names the differences a reader would see.
// Shared by the text-flow fidelity test and the subset matrix.

export const PX_TO_PT = 0.75;
export const X_TOLERANCE_PT = 3;
export const Y_TOLERANCE_PT = 4;
export const SPACING_TOLERANCE_PT = 1;
export const EMU_PER_PT = 12700;
// Office may set a line this much narrower than the browser before a word it pulls
// up counts: LibreOffice lays Arial out 0.55 % narrower than its design widths,
// which Chrome keeps (measured 2026-09-25, same font file on both sides).
export const REWRAP_TOLERANCE = 0.01;

export const normalize = (word) => word.toLocaleLowerCase('de');

// Headless LibreOffice on macOS finds its fonts through the fontconfig it ships,
// which knows only the fonts bundled with it: Arial became Liberation Sans and
// Helvetica Linux Libertine, so every comparison measured a substitute against
// the browser's real font. Pointing fontconfig at the system's font directories
// gives Office the fonts the browser used.
const MACOS_FONT_DIRS = [
  '/System/Library/Fonts',
  '/System/Library/Fonts/Supplemental',
  '/Library/Fonts',
  path.join(os.homedir(), 'Library/Fonts'),
];

/** Converts a PPTX to PDF with LibreOffice in a profile of its own, using the fonts the browser has. */
export function convertToPdf(pptxPath, outputDir, profilePath = path.join(outputDir, 'libreoffice-profile')) {
  const env = { ...process.env };
  if (process.platform === 'darwin') {
    const config = path.join(outputDir, 'fonts.conf');
    writeFileSync(
      config,
      [
        '<?xml version="1.0"?>',
        '<fontconfig>',
        ...MACOS_FONT_DIRS.map((dir) => `  <dir>${dir}</dir>`),
        `  <cachedir>${path.join(outputDir, 'fontconfig-cache')}</cachedir>`,
        '</fontconfig>',
        '',
      ].join('\n')
    );
    env.FONTCONFIG_FILE = config;
  }
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
    { stdio: 'pipe', env }
  );
}

export function groupLines(words) {
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

export async function measureBrowserPages(executablePath, fixture, canvas) {
  const browser = await puppeteer.launch({ executablePath, headless: true, args: ['--no-sandbox'] });
  try {
    const page = await browser.newPage();
    await page.setViewport(canvas);
    await page.goto(pathToFileURL(fixture).href, { waitUntil: 'load' });
    await page.evaluate(() => document.fonts.ready);
    const pages = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.slide'), (slide) => {
        const origin = slide.getBoundingClientRect();
        const words = [];
        // The block a word's lines break in, and the content box they break against.
        const blocks = new Map();
        const blockOf = (element) => {
          let block = element;
          while (block !== slide && getComputedStyle(block).display.startsWith('inline')) block = block.parentElement;
          if (!blocks.has(block)) {
            const rect = block.getBoundingClientRect();
            const style = getComputedStyle(block);
            blocks.set(block, {
              block: blocks.size,
              blockLeft: rect.left - origin.left + parseFloat(style.borderLeftWidth) + parseFloat(style.paddingLeft),
              blockRight:
                rect.right - origin.left - parseFloat(style.borderRightWidth) - parseFloat(style.paddingRight),
            });
          }
          return blocks.get(block);
        };
        const walker = document.createTreeWalker(slide, NodeFilter.SHOW_TEXT);
        let node;
        while ((node = walker.nextNode())) {
          const probe = node.parentElement.closest('[data-probe]')?.dataset.probe ?? null;
          const block = blockOf(node.parentElement);
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
              ...block,
              x: rect.left - origin.left,
              xMax: rect.right - origin.left,
              top: rect.top - origin.top,
              bottom: rect.bottom - origin.top,
            });
          }
        }
        // A slide can be a probe itself, e.g. when its background is what is measured.
        const probeElements = [
          ...(slide.matches('[data-probe]') ? [slide] : []),
          ...slide.querySelectorAll('[data-probe]'),
        ];
        // Every element of a probe that puts paint on the slide itself — a fill, a
        // visible border side, a shadow, a picture. Each is compared in its own box,
        // so a small shape is not diluted by the probe's empty space.
        const visible = (color) => {
          const alpha = color.match(/rgba?\([^)]*[,/]\s*([\d.]+%?)\s*\)/);
          return color !== 'transparent' && (!alpha || parseFloat(alpha[1]) > 0);
        };
        const paints = (element) => {
          const style = getComputedStyle(element);
          if (style.visibility === 'hidden') return false;
          if (element.matches('img, svg')) return true;
          return (
            visible(style.backgroundColor) ||
            style.backgroundImage !== 'none' ||
            style.boxShadow !== 'none' ||
            ['Top', 'Right', 'Bottom', 'Left'].some(
              (side) =>
                parseFloat(style[`border${side}Width`]) > 0 &&
                style[`border${side}Style`] !== 'none' &&
                visible(style[`border${side}Color`])
            )
          );
        };
        const labelOf = (element) =>
          element.tagName.toLowerCase() + (element.dataset.semanticId ? `[${element.dataset.semanticId}]` : '');
        const probes = probeElements.map((element) => {
          const box = element.getBoundingClientRect();
          const painted = [element, ...element.querySelectorAll('*')]
            .filter((candidate) => !candidate.parentElement?.closest('svg') && paints(candidate))
            .map((candidate, index) => {
              const rect = candidate.getBoundingClientRect();
              return {
                label: `${index + 1}:${labelOf(candidate)}`,
                top: rect.top - origin.top,
                bottom: rect.bottom - origin.top,
                left: rect.left - origin.left,
                right: rect.right - origin.left,
              };
            })
            .filter((rect) => rect.right > rect.left && rect.bottom > rect.top);
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
            // margins lie outside the box and therefore count as zero here. A slide
            // has a fixed height, so its leftover space is not paragraph spacing.
            leadingSpace: element === slide ? null : first ? first.top - contentTop : 0,
            trailingSpace: element === slide ? null : last ? contentBottom - last.bottom : 0,
            painted,
          };
        });
        return { words, probes };
      })
    );
    const slideIds = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.slide'), (slide) => slide.dataset.slideId ?? null)
    );
    const toPt = (value) => value * PX_TO_PT;
    const measured = pages.map((page) => ({
      words: page.words.map((word) => ({
        ...word,
        blockLeft: toPt(word.blockLeft),
        blockRight: toPt(word.blockRight),
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
        leadingSpace: probe.leadingSpace === null ? null : toPt(probe.leadingSpace),
        trailingSpace: probe.trailingSpace === null ? null : toPt(probe.trailingSpace),
        painted: probe.painted.map((rect) => ({
          ...rect,
          top: toPt(rect.top),
          bottom: toPt(rect.bottom),
          left: toPt(rect.left),
          right: toPt(rect.right),
        })),
      })),
    }));
    return { pages: measured, slideIds };
  } finally {
    await browser.close();
  }
}

export function parseOfficeWords(bboxHtml) {
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
export function assignWords(browserWords, officeWords) {
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

export function textFrameParagraphs(xml, marker) {
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
 * The first line break per text block where Office parts from the browser; the
 * later ones in the block follow from it. Office adding a break is always one. A
 * word Office pulls up to the line before is one unless the browser broke there
 * for width alone and missed by no more than REWRAP_TOLERANCE of the line — then
 * it is the narrower metrics of the Office renderer, not the converter.
 * Compared on the paired words only: a lost word is `missing-word`'s finding.
 */
function rewrapDetails(pairs) {
  const details = [];
  const byBlock = new Map();
  for (const word of pairs) byBlock.set(word.browser.block, [...(byBlock.get(word.browser.block) ?? []), word]);
  const opens = (lines) => new Set(lines.map((line) => line.words[0].browserIndex));
  for (const words of byBlock.values()) {
    const inBrowser = opens(groupLines(words.map((word) => ({ ...word.browser, browserIndex: word.browserIndex }))));
    const inOffice = opens(groupLines(words));
    const ordered = [...words].sort((a, b) => a.browserIndex - b.browserIndex);
    const index = ordered.findIndex((word) => inBrowser.has(word.browserIndex) !== inOffice.has(word.browserIndex));
    if (index < 1) continue;
    const [previous, word] = [ordered[index - 1], ordered[index]];
    if (inOffice.has(word.browserIndex)) {
      details.push(`Office breaks the line before "${word.text}", the browser does not`);
      continue;
    }
    const gaps = ordered
      .slice(1)
      .map((next, position) => [ordered[position], next])
      .filter(([left, right]) => !inBrowser.has(right.browserIndex) && right.browser.x > left.browser.xMax)
      .map(([left, right]) => right.browser.x - left.browser.xMax)
      .sort((a, b) => a - b);
    const gap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 0.25 * (word.browser.bottom - word.browser.top);
    const width = word.browser.blockRight - word.browser.blockLeft;
    const missing = previous.browser.xMax + gap + (word.browser.xMax - word.browser.x) - word.browser.blockRight;
    if (missing > -gap / 2 && missing <= REWRAP_TOLERANCE * width) continue;
    details.push(
      `Office pulls "${word.text}" up to the line before; the browser line lacked ${missing.toFixed(1)} pt of ${width.toFixed(0)} pt`
    );
  }
  return details;
}

/**
 * Differences a reader would see between the browser and Office, as
 * `{ probe, kind, detail }`. Office has to break every line where the browser
 * broke it: a line more or less changes the height of the text and says the
 * box's width did not carry over.
 */
export function analyzePage(browserPage, officeWords, xml, rasterized = {}) {
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
    if (word.probe in rasterized) return;
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
  for (const probe of browserPage.probes.filter((candidate) => !(candidate.name in rasterized))) {
    const pairs = browserPage.words
      .map((word, index) => ({ word, index }))
      .filter(({ word, index }) => word.probe === probe.name && officeFor.has(index))
      .map(({ word, index }) => ({
        ...officeWords[officeFor.get(index)],
        browserIndex: index,
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

    for (const detail of rewrapDetails(pairs)) add(probe.name, 'rewrap', detail);

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
    if (paragraphs?.length && probe.leadingSpace !== null) {
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
  const stackable = browserPage.probes.filter((candidate) => !(candidate.name in rasterized));
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
