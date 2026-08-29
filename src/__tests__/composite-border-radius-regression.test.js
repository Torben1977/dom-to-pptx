import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';
import { generateCompositeBorderSVG, resolveCssCornerRadii } from '../utils.js';

function decodeSvg(dataUrl) {
  return atob(dataUrl.slice(dataUrl.indexOf(',') + 1));
}

const emptySide = { width: 0, style: 'none', color: null, opacity: 0 };

describe('elliptical composite border radii', () => {
  it('resolves percentage corner radii against the correct box axes', () => {
    expect(resolveCssCornerRadii({
      borderTopLeftRadius: '50%',
      borderTopRightRadius: '25% 40%',
      borderBottomRightRadius: '12px 18px',
      borderBottomLeftRadius: '0px',
    }, 420, 620)).toEqual({
      tl: { x: 210, y: 310 },
      tr: { x: 105, y: 248 },
      br: { x: 12, y: 18 },
      bl: { x: 0, y: 0 },
    });
  });

  it('draws a one-sided border along an elliptical rounded contour', () => {
    const radii = {
      tl: { x: 210, y: 310 },
      tr: { x: 210, y: 310 },
      br: { x: 210, y: 310 },
      bl: { x: 210, y: 310 },
    };
    const svg = decodeSvg(generateCompositeBorderSVG(420, 620, radii, {
      top: emptySide,
      right: emptySide,
      bottom: emptySide,
      left: { width: 8, style: 'solid', color: '096E66', opacity: 1 },
    }));

    expect(svg).toContain('stroke="#096E66"');
    expect(svg).toMatch(/A\s+206\s+306\s+0\s+0\s+0/);
    expect(svg).not.toContain('<rect x="0" y="0" width="8" height="620"');
  });

  it('preserves the percentage-radius arc through the real browser-to-PPTX path', async () => {
    const html = `
      <!doctype html>
      <html>
        <body style="margin:0">
          <section class="slide" style="position:relative;width:1920px;height:1080px;background:#fff">
            <div class="arc"></div>
          </section>
          <style>
            .arc {
              position: absolute;
              left: 1220px;
              top: 180px;
              width: 420px;
              height: 620px;
              box-sizing: border-box;
              border-left: 8px solid #096E66;
              border-radius: 50%;
              transform: rotate(10deg);
            }
          </style>
        </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5 },
      autoEmbedFonts: false,
    });
    const zip = await JSZip.loadAsync(buffer);
    const mediaNames = Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/') && name.endsWith('.svg'));
    const media = await Promise.all(mediaNames.map((name) => zip.file(name).async('string')));
    const arcSvg = media.find((svg) => svg.includes('#096E66'));

    expect(arcSvg).toBeDefined();
    expect(arcSvg).toMatch(/A\s+206\s+306\s+0\s+0\s+0/);
    expect(arcSvg).not.toContain('<rect x="0" y="0" width="8" height="620"');
  }, 40000);
});
