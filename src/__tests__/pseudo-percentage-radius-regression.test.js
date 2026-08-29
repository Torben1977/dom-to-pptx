import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

describe('pseudo-element percentage radii', () => {
  it('maps a 50%-radius rectangular pseudo-element to an ellipse', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
        <style>
          body { margin: 0; }
          .slide { position: relative; width: 960px; height: 540px; overflow: hidden; background: #eef8f3; }
          .slide::after { content: ""; position: absolute; left: 300px; top: 160px;
            width: 240px; height: 160px; border-radius: 50%; background: #b9d9ca; }
        </style>
      </head>
      <body><section class="slide"></section></body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 10, height: 5.625, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const fillIndex = xml.indexOf('<a:srgbClr val="B9D9CA"');
    expect(fillIndex).toBeGreaterThan(-1);
    const shapeStart = xml.lastIndexOf('<p:sp>', fillIndex);
    const shapeEnd = xml.indexOf('</p:sp>', fillIndex);
    const shape = xml.slice(shapeStart, shapeEnd);

    expect(shape).toContain('<a:prstGeom prst="ellipse">');
  }, 40_000);
});
