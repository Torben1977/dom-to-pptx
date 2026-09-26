import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import puppeteer from 'puppeteer';
import { exportHtmlToPptx } from '../node-exporter.js';

// A line's extra leading lands in different places: CSS splits it evenly above
// and below the text, PowerPoint puts about three quarters above, LibreOffice
// all of it. A single-line frame has no line pitch to keep, so the converter
// drops its line spacing: at single spacing both renderers seat the baseline
// about 0.966 em below the frame top (PowerPoint 0.94, LibreOffice 0.99), so the
// frame goes that far above the browser's baseline. A paragraph of several
// lines keeps its spacing, and a loose one is lifted by what PowerPoint adds
// above its first line: 0.234 × (line height − 1.46 × font size), at 20pt/40pt
// 2.527 pt.
const SHIFT_PT = 0.234 * (40 - 1.46 * 20);
const BASELINE_EM = 0.966;
const EMU_PER_PT = 12700;
const EMU_PER_PX = 914_400 / 96;

const slide = (body) => `<!doctype html><html><head><style>
    * { box-sizing: border-box; margin: 0; }
    .slide { position: relative; width: 1280px; height: 720px; background: white; }
    .text { position: absolute; left: 100px; white-space: nowrap; font-family: Arial, sans-serif; }
  </style></head><body><section class="slide">${body}</section></body></html>`;

async function exportXml(html) {
  const buffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
  });
  return (await JSZip.loadAsync(buffer)).file('ppt/slides/slide1.xml').async('string');
}

/** The browser's baseline and font size of each `.text` element, in CSS px, by its text. */
async function browserBaselines(html) {
  const browser = await puppeteer.launch({
    executablePath: await puppeteer.executablePath(),
    headless: true,
    args: ['--no-sandbox'],
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 720 });
    await page.setContent(html);
    return await page.evaluate(() =>
      Object.fromEntries(
        Array.from(document.querySelectorAll('.text'), (element) => {
          // Without a wrap opportunity the marker stays on the line even where
          // the text is wider than its box.
          const whiteSpace = element.style.whiteSpace;
          element.style.whiteSpace = 'nowrap';
          const marker = document.createElement('span');
          marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
          element.appendChild(marker);
          const baseline = marker.getBoundingClientRect().bottom;
          marker.remove();
          element.style.whiteSpace = whiteSpace;
          return [element.textContent.trim(), { baseline, fontPx: parseFloat(getComputedStyle(element).fontSize) }];
        })
      )
    );
  } finally {
    await browser.close();
  }
}

function shapeFor(xml, text) {
  const index = xml.indexOf(`<a:t>${text}</a:t>`);
  expect(index, `text '${text}'`).toBeGreaterThan(-1);
  return xml.slice(xml.lastIndexOf('<p:sp>', index), xml.indexOf('</p:sp>', index));
}

const topPx = (shape) => Number(shape.match(/<a:off x="\d+" y="(\d+)"/)[1]) / EMU_PER_PX;
const topInsetPt = (shape) => Number(shape.match(/tIns="(\d+)"/)?.[1] ?? 0) / EMU_PER_PT;

describe("PowerPoint's leading", () => {
  it("seats a single-line frame's baseline on the browser's, at single line spacing", async () => {
    const html = slide(
      '<p class="text" style="top: 100px; font-size: 20pt; line-height: 40pt">Lockere Zeile</p>' +
        '<p class="text" style="top: 250px; font-size: 20pt; line-height: 29.2pt">Neutrale Zeile</p>' +
        '<p class="text" style="top: 400px; font-size: 20pt; line-height: 24pt">Knappe Zeile</p>' +
        '<p class="text" style="top: 550px; font-size: 20pt">Normale Zeile</p>'
    );
    const [xml, lines] = await Promise.all([exportXml(html), browserBaselines(html)]);

    for (const text of ['Lockere Zeile', 'Neutrale Zeile', 'Knappe Zeile', 'Normale Zeile']) {
      const shape = shapeFor(xml, text);
      const { baseline, fontPx } = lines[text];
      expect(topPx(shape), text).toBeCloseTo(baseline - BASELINE_EM * fontPx, 0);
      expect(shape, text).toContain('<a:spcPct val="100000"/>');
      expect(shape, text).not.toContain('<a:spcPts');
    }
  });

  it('seats the baseline of a word wider than its box', async () => {
    const html = slide(
      '<p class="text" style="top: 100px; width: 60px; white-space: normal; font-size: 42pt; line-height: 48pt">01</p>'
    );
    const [xml, lines] = await Promise.all([exportXml(html), browserBaselines(html)]);
    const { baseline, fontPx } = lines['01'];

    expect(topPx(shapeFor(xml, '01'))).toBeCloseTo(baseline - BASELINE_EM * fontPx, 0);
  });

  it("keeps a painted single-line frame in place and seats its baseline on the browser's through the inset", async () => {
    const html = slide(
      '<p class="text" style="top: 100px; padding: 10px; background: #EEF1F5; font-size: 20pt; line-height: 40pt">Kasten</p>'
    );
    const [xml, lines] = await Promise.all([exportXml(html), browserBaselines(html)]);
    const shape = shapeFor(xml, 'Kasten');
    const { baseline, fontPx } = lines.Kasten;

    expect(topPx(shape)).toBeCloseTo(100, 1);
    expect(topInsetPt(shape)).toBeCloseTo((baseline - BASELINE_EM * fontPx - 100) * 0.75, 0);
  });

  it('lifts a loose paragraph of several lines by the leading PowerPoint adds above its first line', async () => {
    const xml = await exportXml(
      slide(
        '<p class="text" style="top: 100px; width: 260px; white-space: normal; font-size: 20pt; line-height: 40pt">' +
          'Ein Absatz über mehrere Zeilen</p>'
      )
    );
    const shape = shapeFor(xml, 'Ein Absatz über mehrere Zeilen');

    expect(topPx(shape)).toBeCloseTo(100 - SHIFT_PT / 0.75, 1);
    expect(shape).toContain('<a:spcPts val="4000"/>');
  });

  it('lifts the text of a table cell within its top margin', async () => {
    const xml = await exportXml(
      slide(
        '<table style="position: absolute; left: 100px; top: 100px; border-collapse: collapse; font-family: Arial, sans-serif">' +
          '<tr><td style="padding: 10px; font-size: 20pt; line-height: 40pt">Zelle</td></tr></table>'
      )
    );
    const cell = Array.from(xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g), (match) => match[0]).find((tc) =>
      tc.includes('<a:t>Zelle</a:t>')
    );
    const marginTopPt = Number(cell.match(/<a:tcPr[^>]*\bmarT="(\d+)"/)[1]) / EMU_PER_PT;

    expect(marginTopPt).toBeCloseTo(7.5 - SHIFT_PT, 1);
  });
});
