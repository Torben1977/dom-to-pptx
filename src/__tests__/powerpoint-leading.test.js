import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

// PowerPoint sets a line's extra leading mostly above its text; CSS splits it
// evenly. Measured, the first line lands 0.234 × (line height − 1.46 × font
// size) below the browser's, so the converter lifts lines looser than 1.46 by
// that much: a frame that paints nothing as a whole, a painted frame and a
// table cell through their top inset. At 20pt/40pt that is 0.234 × (40 − 29.2)
// = 2.527 pt. Tighter lines stay put, where LibreOffice would suffer.
const SHIFT_PT = 0.234 * (40 - 1.46 * 20);
const EMU_PER_PT = 12700;
const EMU_PER_PX = 914_400 / 96;

async function exportSlide(body) {
  const html = `<!doctype html><html><head><style>
      * { box-sizing: border-box; margin: 0; }
      .slide { position: relative; width: 1280px; height: 720px; background: white; }
      .text { position: absolute; left: 100px; white-space: nowrap; font-family: Arial, sans-serif; }
    </style></head><body><section class="slide">${body}</section></body></html>`;
  const buffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
  });
  return (await JSZip.loadAsync(buffer)).file('ppt/slides/slide1.xml').async('string');
}

function shapeFor(xml, text) {
  const index = xml.indexOf(`<a:t>${text}</a:t>`);
  expect(index, `text '${text}'`).toBeGreaterThan(-1);
  return xml.slice(xml.lastIndexOf('<p:sp>', index), xml.indexOf('</p:sp>', index));
}

const topPx = (shape) => Number(shape.match(/<a:off x="\d+" y="(\d+)"/)[1]) / EMU_PER_PX;
const topInsetPt = (shape) => Number(shape.match(/tIns="(\d+)"/)?.[1] ?? 0) / EMU_PER_PT;

describe("PowerPoint's leading", () => {
  it('lifts an unpainted frame by the leading PowerPoint adds above a loose first line', async () => {
    const xml = await exportSlide(
      '<p class="text" style="top: 100px; font-size: 20pt; line-height: 40pt">Lockere Zeile</p>' +
        '<p class="text" style="top: 300px; font-size: 20pt; line-height: 29.2pt">Neutrale Zeile</p>' +
        '<p class="text" style="top: 400px; font-size: 20pt; line-height: 24pt">Knappe Zeile</p>' +
        '<p class="text" style="top: 500px; font-size: 20pt">Normale Zeile</p>'
    );

    expect(topPx(shapeFor(xml, 'Lockere Zeile'))).toBeCloseTo(100 - SHIFT_PT / 0.75, 1);
    expect(topPx(shapeFor(xml, 'Neutrale Zeile'))).toBeCloseTo(300, 1);
    // A tight line stays: the model alone would have moved it about 1.6 px down.
    const tightTop = topPx(shapeFor(xml, 'Knappe Zeile'));
    expect(tightTop).toBeGreaterThan(398.5);
    expect(tightTop).toBeLessThan(400.5);
    // Without a line height there is no exact line spacing to correct for; the
    // model would have moved this line about 1.9 px down.
    const normalTop = topPx(shapeFor(xml, 'Normale Zeile'));
    expect(normalTop).toBeGreaterThan(498.5);
    expect(normalTop).toBeLessThan(500.5);
  });

  it('keeps a painted frame in place and moves its text within the top inset', async () => {
    const xml = await exportSlide(
      '<p class="text" style="top: 100px; padding: 10px; background: #EEF1F5; font-size: 20pt; line-height: 40pt">Kasten</p>'
    );
    const shape = shapeFor(xml, 'Kasten');

    expect(topPx(shape)).toBeCloseTo(100, 1);
    expect(topInsetPt(shape)).toBeCloseTo(7.5 - SHIFT_PT, 1);
  });

  it('moves the text of a table cell within its top margin', async () => {
    const xml = await exportSlide(
      '<table style="position: absolute; left: 100px; top: 100px; border-collapse: collapse; font-family: Arial, sans-serif">' +
        '<tr><td style="padding: 10px; font-size: 20pt; line-height: 40pt">Zelle</td></tr></table>'
    );
    const cell = Array.from(xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g), (match) => match[0]).find((tc) =>
      tc.includes('<a:t>Zelle</a:t>')
    );
    const marginTopPt = Number(cell.match(/<a:tcPr[^>]*\bmarT="(\d+)"/)[1]) / EMU_PER_PT;

    expect(marginTopPt).toBeCloseTo(7.5 - SHIFT_PT, 1);
  });
});
