import puppeteer from 'puppeteer';
import { pathToFileURL } from 'node:url';

// Measures where every word and every probe box lies in the browser and where
// Office puts the same words, and names the differences a reader would see.
// Shared by the text-flow fidelity test and the subset matrix.

export const PX_TO_PT = 0.75;
export const X_TOLERANCE_PT = 3;
export const Y_TOLERANCE_PT = 4;
export const SPACING_TOLERANCE_PT = 1;
export const EMU_PER_PT = 12700;

export const normalize = (word) => word.toLocaleLowerCase('de');

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
        // A slide can be a probe itself, e.g. when its background is what is measured.
        const probeElements = [
          ...(slide.matches('[data-probe]') ? [slide] : []),
          ...slide.querySelectorAll('[data-probe]'),
        ];
        const probes = probeElements.map((element) => {
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
            // margins lie outside the box and therefore count as zero here. A slide
            // has a fixed height, so its leftover space is not paragraph spacing.
            leadingSpace: element === slide ? null : first ? first.top - contentTop : 0,
            trailingSpace: element === slide ? null : last ? contentBottom - last.bottom : 0,
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
 * Differences a reader would see between the browser and Office, as
 * `{ probe, kind, detail }`. Wrapping at a different word is not one by itself.
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
