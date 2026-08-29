import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

function shapeFor(xml, text) {
  const textIndex = xml.indexOf(`<a:t>${text}</a:t>`);
  expect(textIndex, `text '${text}'`).toBeGreaterThan(-1);
  const shapeStart = xml.lastIndexOf('<p:sp>', textIndex);
  const shapeEnd = xml.indexOf('</p:sp>', textIndex);
  return xml.slice(shapeStart, shapeEnd);
}

function horizontalGeometry(shape) {
  const x = Number(shape.match(/<a:off x="(\d+)"/)?.[1]);
  const width = Number(shape.match(/<a:ext cx="(\d+)"/)?.[1]);
  return { x, width, right: x + width };
}

describe('browser-resolved float text flow', () => {
  it('places editable text in the browser line region beside left and right floats', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
        <style>
          * { box-sizing: border-box; }
          body { margin: 0; }
          .slide { position: relative; width: 1920px; height: 1080px; overflow: hidden; background: white; }
          .card { position: absolute; width: 700px; height: 220px; padding: 28px 32px; background: #f7f9fc; }
          .left-card { left: 120px; top: 160px; }
          .right-card { left: 980px; top: 160px; }
          .badge { width: 92px; height: 72px; font: 700 48px/72px Arial; color: #065dc9; }
          .left-card .badge { float: left; margin-right: 24px; }
          .right-card .badge { float: right; margin-left: 24px; }
          h3 { margin: 0; font: 700 28px/36px Arial; color: #0e1b2c; white-space: nowrap; }
        </style>
      </head>
      <body>
        <section class="slide">
          <div class="card left-card"><div class="badge">01</div><h3>Left-float heading</h3></div>
          <div class="card right-card"><div class="badge">02</div><h3>Right-float heading</h3></div>
        </section>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    const leftBadge = horizontalGeometry(shapeFor(xml, '01'));
    const leftHeading = horizontalGeometry(shapeFor(xml, 'Left-float heading'));
    expect(leftHeading.x).toBeGreaterThanOrEqual(leftBadge.right);

    const rightBadge = horizontalGeometry(shapeFor(xml, '02'));
    const rightHeading = horizontalGeometry(shapeFor(xml, 'Right-float heading'));
    expect(rightHeading.right).toBeLessThanOrEqual(rightBadge.x);
  }, 40_000);

  it('keeps the representative floated card number out of its heading', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
        <style>
          * { box-sizing: border-box; }
          body { margin: 0; }
          .slide { position: relative; width: 1920px; height: 1080px; overflow: hidden; background: white; }
          .cause { position: absolute; left: 120px; top: 200px; width: 780px; height: 235px;
            padding: 28px 32px; border-top: 7px solid #065dc9; background: #f7f9fc; border-radius: 8px; }
          .cause .n { float: left; margin-right: 22px; color: #065dc9; font: 700 48px/50px Arial; }
          .cause h3 { margin: 0 0 12px; color: #0e1b2c; font: 700 23px/29px Arial; }
          .cause p { margin: 0; color: #4b5563; font: 400 18px/27px Arial; }
        </style>
      </head>
      <body>
        <section class="slide">
          <div class="cause"><div class="n">01</div><h3>Unverbundene Piloten</h3><p>32 von 46 Fällen starteten aus einzelnen Funktionen.</p></div>
        </section>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const number = horizontalGeometry(shapeFor(xml, '01'));
    const heading = horizontalGeometry(shapeFor(xml, 'Unverbundene Piloten'));

    expect(heading.x).toBeGreaterThanOrEqual(number.right);
    expect(shapeFor(xml, 'Unverbundene Piloten')).not.toContain('<a:normAutofit');
  }, 40_000);
});
