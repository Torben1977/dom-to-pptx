import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

const html = `
  <!doctype html>
  <html>
    <head>
      <style>
        * { box-sizing: border-box; }
        body { margin: 0; }
        .slide {
          position: relative;
          width: 960px;
          height: 540px;
          overflow: hidden;
          background: #fff;
        }
        .clip {
          position: absolute;
          left: 100px;
          top: 80px;
          width: 200px;
          height: 120px;
          overflow: hidden;
          background: #eef2ff;
        }
        .text-child {
          position: absolute;
          left: 150px;
          top: 30px;
          width: 140px;
          height: 44px;
          background: #f79009;
          color: #fff;
          font: 700 20px/1.2 Arial, sans-serif;
        }
        .solid-child {
          position: absolute;
          left: 150px;
          top: 30px;
          width: 140px;
          height: 44px;
          background: #d92d20;
        }
        .transform-child {
          position: absolute;
          left: 150px;
          top: 30px;
          width: 140px;
          height: 44px;
          transform: rotate(8deg);
          background: #6f42c1;
        }
      </style>
    </head>
    <body>
      <section class="slide" id="text-slide">
        <div class="clip" id="text-clip"><div class="text-child">CLIPPED_TEXT</div></div>
      </section>
      <section class="slide" id="solid-slide">
        <div class="clip" id="solid-clip"><div class="solid-child"></div></div>
      </section>
      <section class="slide" id="transform-slide">
        <div class="clip" id="transform-clip"><div class="transform-child"></div></div>
      </section>
    </body>
  </html>
`;

const pptxOptions = {
  width: 10,
  height: 5.625,
  autoEmbedFonts: false,
};

describe('overflow clipping boundary policy', () => {
  it('rejects clipped editable text with an actionable structured finding', async () => {
    await expect(
      exportHtmlToPptx(html, {
        selector: '#text-slide',
        pptxOptions: { ...pptxOptions, boundaryPolicy: 'error' },
      })
    ).rejects.toThrow(/DOM_TO_PPTX_UNSUPPORTED_BOUNDARY.*overflow-clipping.*#text-clip.*\.text-child/s);
  });

  it('keeps simple axis-aligned solid clipping native and editable in strict mode', async () => {
    const buffer = await exportHtmlToPptx(html, {
      selector: '#solid-slide',
      pptxOptions: { ...pptxOptions, boundaryPolicy: 'error' },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    expect(xml).toContain('D92D20');
    expect(xml).not.toContain('<p:pic>');
  });

  it('rasterizes only the affected clipping subtree in fidelity mode', async () => {
    const buffer = await exportHtmlToPptx(html, {
      selector: '#text-slide',
      pptxOptions: { ...pptxOptions, boundaryPolicy: 'rasterize' },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    expect(xml).toContain('<p:pic>');
    expect(xml).not.toContain('CLIPPED_TEXT');
  });

  it('rejects transformed clipped descendants instead of approximating their crop', async () => {
    await expect(
      exportHtmlToPptx(html, {
        selector: '#transform-slide',
        pptxOptions: { ...pptxOptions, boundaryPolicy: 'error' },
      })
    ).rejects.toThrow(/overflow-clipping.*#transform-clip.*\.transform-child/s);
  });

  it('does not classify the slide canvas overflow as an unsupported nested boundary', async () => {
    const buffer = await exportHtmlToPptx(
      '<section class="slide" style="position:relative;width:960px;height:540px;overflow:hidden">OK</section>',
      {
        selector: '.slide',
        pptxOptions: { ...pptxOptions, boundaryPolicy: 'error' },
      }
    );

    expect(buffer.length).toBeGreaterThan(0);
  });
});
