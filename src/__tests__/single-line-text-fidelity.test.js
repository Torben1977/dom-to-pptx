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

function shapeGeometry(shape) {
  const x = Number(shape.match(/<a:off x="(\d+)"/)?.[1]);
  const width = Number(shape.match(/<a:ext cx="(\d+)"/)?.[1]);
  return { x, width, right: x + width };
}

describe('browser single-line fidelity', () => {
  it('keeps natural-width metrics and anonymous flex labels on one PowerPoint line', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
        <style>
          * { box-sizing: border-box; }
          body { margin: 0; }
          .slide { position: relative; width: 1920px; height: 1080px; overflow: hidden; background: white; }
          .metric { position: absolute; left: 120px; top: 180px; width: max-content;
            font: 700 76px/80px Arial, sans-serif; color: #f24726; }
          .bar { position: absolute; left: 120px; top: 360px; height: 44px; width: 620px;
            display: flex; align-items: center; padding-left: 20px; border-radius: 22px;
            background: #24558e; color: white; font: 700 20px/24px Arial, sans-serif; }
          .bar small { margin-left: 14px; font: 400 15px/20px Arial, sans-serif; }
          .cost { position: absolute; left: 980px; top: 180px; width: 720px; height: 131px;
            padding-top: 22px; border-top: 1px solid #d1d5db; }
          .cost .label { display: block; font: 700 14px/20px Arial, sans-serif; }
          .cost strong { display: inline; font: 700 43px/64.5px Arial, sans-serif; }
          .cost .support { display: block; font: 400 16px/23px Arial, sans-serif; }
        </style>
      </head>
      <body>
        <section class="slide">
          <div class="metric">2,6 Mio. €</div>
          <div class="bar">5–10 %<small>12 Monate</small></div>
          <div class="cost"><div class="label">PLATTFORMKOSTEN</div><strong>4,1 Mio. €</strong><div class="support">höher als geplant</div></div>
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

    for (const text of ['2,6 Mio. €', '12 Monate']) {
      const shape = shapeFor(xml, text);
      expect(shape).toContain('wrap="none"');
      expect(shape).not.toContain('<a:normAutofit');
      expect(shape).not.toContain('<a:spAutoFit');
    }

    const primaryMetric = shapeFor(xml, '2,6 Mio. €');
    expect(Number(primaryMetric.match(/<a:ext cx="(\d+)"/)?.[1])).toBeGreaterThan(2_000_000);

    const anonymousBarLabel = shapeFor(xml, '5–10 %');
    expect(anonymousBarLabel).toContain('wrap="none"');
    expect(anonymousBarLabel).not.toContain('<a:normAutofit');
    expect(Number(anonymousBarLabel.match(/<a:ext cx="(\d+)"/)?.[1])).toBeGreaterThan(420_000);
    expect(Number(anonymousBarLabel.match(/<a:ext cx="\d+" cy="(\d+)"/)?.[1])).toBeGreaterThan(150_000);

    const costShape = shapeFor(xml, '4,1 Mio. €');
    expect(costShape).toContain('wrap="none"');
    expect(costShape).not.toContain('<a:normAutofit');
    const width = Number(costShape.match(/<a:ext cx="(\d+)"/)?.[1]);
    const height = Number(costShape.match(/<a:ext cx="\d+" cy="(\d+)"/)?.[1]);
    // The KPI sits in a deliberately narrow column. Preserve its authored
    // single line and font size without inventing a slide-wide text box.
    expect(width).toBeGreaterThan(1_200_000);
    expect(height).toBeGreaterThan(400_000);
  }, 40_000);

  it('honors a normalized browser line-count contract without measuring the frozen DOM again', async () => {
    const html = `
      <!doctype html>
      <html><body style="margin:0">
        <section class="slide" style="position:relative;width:1920px;height:1080px;background:#fff">
          <div data-pptx-rendered-lines="1" style="position:absolute;left:120px;top:120px;width:180px;
            white-space:normal;font:700 32px/38px Arial;color:#102033">
            Fixed browser snapshot contract
          </div>
        </section>
      </body></html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    expect(shapeFor(xml, 'Fixed browser snapshot contract')).toContain('wrap="none"');
  }, 40_000);

  it('clamps anonymous flex text reserve before the next flex item', async () => {
    const html = `
      <!doctype html>
      <html><body style="margin:0">
        <section class="slide" style="position:relative;width:1920px;height:1080px;background:#fff">
          <div style="position:absolute;left:120px;top:120px;width:420px;height:70px;display:flex;
            align-items:center;gap:1px;font:700 36px/44px Arial;color:#102033">ABCDE<span
              style="display:block;width:180px;height:60px;background:#e6eef9">NEIGHBOR</span></div>
        </section>
      </body></html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const anonymous = shapeGeometry(shapeFor(xml, 'ABCDE'));
    const neighbor = shapeGeometry(shapeFor(xml, 'NEIGHBOR'));

    expect(anonymous.right).toBeLessThanOrEqual(neighbor.x);
  }, 40_000);

  it('reserves geometry instead of shrinking a normalized one-line axis label', async () => {
    const html = `
      <!doctype html>
      <html><body style="margin:0">
        <section class="slide" style="position:relative;width:1920px;height:1080px;background:#fff">
          <div data-pptx-rendered-lines="1" style="position:absolute;left:1004.69px;top:436.5px;
            width:232.312px;height:22.5px;overflow:visible;white-space:nowrap;
            font:400 15px/22.5px Arial,sans-serif;color:#6b7280">
            Wirksame, aber teure Einzellösung
          </div>
        </section>
      </body></html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const shape = shapeFor(xml, 'Wirksame, aber teure Einzellösung');
    const width = Number(shape.match(/<a:ext cx="(\d+)"/)?.[1]);

    expect(shape).toContain('wrap="none"');
    expect(shape).not.toContain('<a:normAutofit');
    expect(width).toBeGreaterThan(1_500_000);
  }, 40_000);
});
