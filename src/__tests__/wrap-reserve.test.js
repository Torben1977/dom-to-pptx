import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

// A wrapped text frame exactly as wide as the browser's breaks elsewhere in an
// Office renderer that sets the text a hair wider. It gets the room the
// browser's own breaks allow: the full reserve where every break has room to
// spare, half the shortfall where one break nearly did not happen.

const EMU_PER_PX = 914_400 / 96;

function shapeFor(xml, text) {
  const textIndex = xml.indexOf(text);
  expect(textIndex, `text '${text}'`).toBeGreaterThan(-1);
  return xml.slice(xml.lastIndexOf('<p:sp>', textIndex), xml.indexOf('</p:sp>', textIndex));
}

function geometry(shape) {
  return {
    width: Number(shape.match(/<a:ext cx="(\d+)"/)[1]) / EMU_PER_PX,
    rightInset: Number(shape.match(/rIns="(\d+)"/)?.[1] ?? 0) / EMU_PER_PX,
  };
}

async function exportSlide(body) {
  const html = `<!doctype html><html><head><style>
      * { box-sizing: border-box; }
      body { margin: 0; }
      .slide { position: relative; width: 1280px; height: 720px; overflow: hidden; background: white; }
      .text { position: absolute; left: 80px; margin: 0; font: 700 18px/26px Arial, sans-serif; color: #111827; }
    </style></head><body><section class="slide">${body}</section></body></html>`;
  const buffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
  });
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('ppt/slides/slide1.xml').async('string');
}

describe('wrap reserve of multi-line text frames', () => {
  it('gives a frame as wide as its longest word the full reserve', async () => {
    const xml = await exportSlide(
      '<p class="text" style="top: 80px; width: min-content">standortübergreifend genutzt</p>'
    );

    const { width } = geometry(shapeFor(xml, 'genutzt'));
    // The browser box is the word's width; the frame grows by max(8 px, 3 %).
    expect(width).toBeGreaterThan(180);
    const browserWidth = await widthOf('standortübergreifend');
    expect(width - browserWidth).toBeCloseTo(8, 0);
  });

  it('keeps a break the browser nearly did not make', async () => {
    // 360 px holds everything but "darum." by about one pixel in Chrome.
    const xml = await exportSlide(
      '<p class="text" style="top: 200px; width: 360px; font-weight: 400">Text direkt in der Fläche, ohne Absatz darum.</p>'
    );

    const { width } = geometry(shapeFor(xml, 'Absatz'));
    expect(width - 360).toBeGreaterThanOrEqual(0);
    expect(width - 360).toBeLessThan(1);
  });

  it('grows the text frame of an accent-lined heading and leaves its line as painted', async () => {
    // A top border alone is drawn by a shape of its own, so the text frame may grow. The box
    // is exactly as wide as its word, as a grid column the heading fills.
    const browserWidth = Math.ceil(await widthOf('Wohnberechtigung'));
    const xml = await exportSlide(
      `<p class="text" style="top: 440px; width: ${browserWidth}px; border-top: 2px solid #096E66; padding-top: 12px">` +
        'Wohnberechtigung</p>'
    );

    const { width } = geometry(shapeFor(xml, 'Wohnberechtigung'));
    const line = Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (match) => match[0]).find((shape) =>
      shape.includes('<a:srgbClr val="096E66"/>')
    );
    expect(width - browserWidth).toBeCloseTo(8, 0);
    expect(geometry(line).width).toBeCloseTo(browserWidth, 0);
  });

  it('takes the reserve of a filled card out of its inset and leaves the card as painted', async () => {
    const xml = await exportSlide(
      '<p class="text" style="top: 320px; width: min-content; padding: 10px 12px; background: #eef2ff">' +
        'standortübergreifend genutzt</p>'
    );

    const { width, rightInset } = geometry(shapeFor(xml, 'genutzt'));
    const browserWidth = (await widthOf('standortübergreifend')) + 24;
    expect(width).toBeCloseTo(browserWidth, 0);
    expect(rightInset).toBeCloseTo(12 - 8, 0);
  });
});

// A column the browser squeezes to its longest word leaves that word no room at
// all. The cell cannot outgrow its column, so the reserve comes out of the cell
// inset on the side the text grows towards; the column keeps its width.
describe('wrap reserve of table cells', () => {
  const WORD = 'Ausführungskomplexität';
  const FILLER = 'Niedrig, mittel oder hoch, je nachdem wie viele Standorte beteiligt sind';

  const marginsPx = (cell) => {
    const margin = (side) => Number(cell.match(new RegExp(`<a:tcPr[^>]*\\b${side}="(\\d+)"`))?.[1] ?? 0) / EMU_PER_PX;
    return { left: margin('marL'), right: margin('marR') };
  };

  // Each table is 420 px wide, so its first column shrinks to WORD plus padding.
  const table = (top, cellStyle) =>
    `<table style="position: absolute; left: 80px; top: ${top}px; width: 420px; border-collapse: collapse; font: 700 18px/26px Arial, sans-serif">` +
    `<tr><td style="${cellStyle}">${WORD}</td><td style="padding: 4px 12px">${FILLER}</td></tr></table>`;

  it('takes the reserve out of the inset on the side the text grows towards', async () => {
    const xml = await exportSlide(
      table(40, 'padding: 4px 12px') +
        table(200, 'padding: 4px 12px; text-align: center') +
        table(360, 'padding: 4px 12px; text-align: right') +
        table(520, 'padding: 4px 3px')
    );
    const cells = Array.from(xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g), (match) => match[0]).filter((tc) =>
      tc.includes(`<a:t>${WORD}</a:t>`)
    );
    expect(cells).toHaveLength(4);
    // One line and no break to measure against: the full max(8 px, 3 %).
    const reserve = Math.max(8, 0.03 * (await widthOf(WORD)));
    const [left, center, right, tight] = cells.map(marginsPx);

    expect(left.left).toBeCloseTo(12, 1);
    expect(left.right).toBeCloseTo(12 - reserve, 1);
    expect(center.left).toBeCloseTo(12 - reserve / 2, 1);
    expect(center.right).toBeCloseTo(12 - reserve / 2, 1);
    expect(right.left).toBeCloseTo(12 - reserve, 1);
    expect(right.right).toBeCloseTo(12, 1);
    // An inset smaller than the reserve is used up, never turned negative.
    expect(tight.left).toBeCloseTo(3, 1);
    expect(tight.right).toBe(0);
  });
});

/** The width Chrome gives a word in the test's font, for comparing frames against. */
async function widthOf(word) {
  const { default: puppeteer } = await import('puppeteer');
  const browser = await puppeteer.launch({
    executablePath: await puppeteer.executablePath(),
    headless: true,
    args: ['--no-sandbox'],
  });
  try {
    const page = await browser.newPage();
    await page.setContent(`<span style="font: 700 18px Arial, sans-serif; white-space: nowrap">${word}</span>`);
    return await page.evaluate(() => document.querySelector('span').getBoundingClientRect().width);
  } finally {
    await browser.close();
  }
}
