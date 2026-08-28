import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

const html = `
  <!doctype html>
  <html>
    <head>
      <style>
        body { margin: 0; }
        .slide { width: 960px; height: 540px; position: relative; }
        .stop-a { stop-color: #d92d20; stop-opacity: .4; }
        .stop-b { stop-color: #1764d8; }
        .dash {
          fill: none;
          stroke: #0f766e;
          stroke-width: 8;
          stroke-opacity: .6;
          stroke-dasharray: 10 5;
          stroke-dashoffset: 3;
        }
      </style>
    </head>
    <body>
      <section class="slide">
        <svg width="400" height="200" viewBox="0 0 400 200">
          <defs>
            <linearGradient id="gradient"><stop class="stop-a"></stop><stop class="stop-b" offset="1"></stop></linearGradient>
          </defs>
          <rect width="400" height="200" fill="url(#gradient)"></rect>
          <path class="dash" d="M20 100H380"></path>
        </svg>
      </section>
    </body>
  </html>
`;

let embeddedSvg;

beforeAll(async () => {
  const buffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: {
      width: 10,
      height: 5.625,
      autoEmbedFonts: false,
      svgAsVector: true,
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  const svgPath = Object.keys(zip.files).find((path) => path.startsWith('ppt/media/') && path.endsWith('.svg'));
  expect(svgPath).toBeDefined();
  embeddedSvg = await zip.file(svgPath).async('string');
}, 40_000);

describe('SVG stylesheet fidelity', () => {
  it('freezes class-based gradient and dash presentation into the embedded SVG', () => {
    expect(embeddedSvg).toMatch(/stop-color:\s*(?:rgb\(217,\s*45,\s*32\)|#d92d20)/i);
    expect(embeddedSvg).toMatch(/stop-opacity:\s*0\.4/i);
    expect(embeddedSvg).toMatch(/stroke-opacity:\s*0\.6/i);
    expect(embeddedSvg).toMatch(/stroke-dasharray:\s*10px?[,]?\s+5px?/i);
    expect(embeddedSvg).toMatch(/stroke-dashoffset:\s*3px?/i);
  });
});
