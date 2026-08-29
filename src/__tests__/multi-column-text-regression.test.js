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

describe('simple native CSS multi-column text', () => {
  it('keeps direct block items as separate editable column text boxes', async () => {
    const html = `
      <!doctype html>
      <html><body style="margin:0">
        <section class="slide" style="position:relative;width:960px;height:540px;background:#fff">
          <div class="columns" style="position:absolute;left:80px;top:80px;width:800px;height:300px;
            column-count:2;column-gap:40px;font:400 24px/32px Arial,sans-serif;color:#102033">
            <p style="margin:0 0 20px">LEFT_COLUMN_TEXT</p>
            <p style="margin:0;break-before:column">RIGHT_COLUMN_TEXT</p>
          </div>
        </section>
      </body></html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 10, height: 5.625, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const left = shapeFor(xml, 'LEFT_COLUMN_TEXT');
    const right = shapeFor(xml, 'RIGHT_COLUMN_TEXT');
    const leftX = Number(left.match(/<a:off x="(\d+)"/)?.[1]);
    const rightX = Number(right.match(/<a:off x="(\d+)"/)?.[1]);

    expect(left).not.toContain('RIGHT_COLUMN_TEXT');
    expect(right).not.toContain('LEFT_COLUMN_TEXT');
    expect(rightX - leftX).toBeGreaterThan(3_000_000);
  }, 40_000);

  it('rejects fragmented multi-column text in controlled editable mode', async () => {
    const html = `
      <section class="slide" data-slide-id="columns-slide"
        style="position:relative;width:960px;height:540px;background:#fff">
        <div data-semantic-id="fragmented-copy" style="position:absolute;left:80px;top:80px;
          width:800px;height:300px;column-count:2;column-gap:40px;font:24px/32px Arial">
          Direct text that the browser is free to balance and fragment between both columns.
        </div>
      </section>`;

    await expect(exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 10, height: 5.625, autoEmbedFonts: false, boundaryPolicy: 'error' },
    })).rejects.toThrow(/multi-column-fragmentation.*columns-slide.*fragmented-copy.*direct-text/s);
  }, 40_000);

  it('rasterizes only the fragmented multi-column subtree in fidelity mode', async () => {
    const html = `
      <section class="slide" style="position:relative;width:960px;height:540px;background:#fff">
        <div style="position:absolute;left:80px;top:80px;width:800px;height:300px;
          column-count:2;column-gap:40px;font:24px/32px Arial">
          Direct text that the browser is free to balance and fragment between both columns.
        </div>
      </section>`;
    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 10, height: 5.625, autoEmbedFonts: false, boundaryPolicy: 'rasterize' },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    expect(xml).toContain('<p:pic>');
    expect(xml).not.toContain('Direct text that the browser');
  }, 40_000);
});
