import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import puppeteer from 'puppeteer';
import { exportHtmlToPptx } from '../node-exporter.js';

// A source line whose overflow clip cuts its 12 px line to a 10 px grid row is
// rasterized under boundaryPolicy 'rasterize'. The capture treated every <span>
// as an icon -- flex-centred, FontAwesome forced, a serif as fallback -- so in
// the Klimametrik board the picture came out centred and in the wrong face,
// while the browser set the line left-aligned in Arial.
const HTML = `<!doctype html><html><head><style>
    * { box-sizing: border-box; margin: 0; }
    .slide { position: relative; width: 1280px; height: 720px; background: white; }
    .foot { position: absolute; left: 64px; top: 600px; width: 1000px; display: grid; grid-template-rows: 10px; font: 7pt/9pt Arial, sans-serif; color: #111827; }
    .src { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  </style></head><body><section class="slide"><div class="foot"><span class="src">Quelle: KlimaMetrik Fakten · Kapazitätsbandbreiten; STRAL · Szenarien</span></div></section></body></html>`;

describe('rasterized text span', () => {
  it('keeps the text of a rasterized span where and as wide as the browser set it', async () => {
    const findings = [];
    const buffer = await exportHtmlToPptx(HTML, {
      selector: '.slide',
      pptxOptions: {
        width: 13.333333,
        height: 7.5,
        autoEmbedFonts: false,
        boundaryPolicy: 'rasterize',
        onBoundaryFindings: (found) => findings.push(...found),
      },
    });
    expect(findings.map((finding) => finding.type)).toContain('overflow-clipping');

    const zip = await JSZip.loadAsync(buffer);
    const imagePath = Object.keys(zip.files).find((name) => /^ppt\/media\/.+\.png$/.test(name));
    expect(imagePath).toBeDefined();
    const png = (await zip.file(imagePath).async('nodebuffer')).toString('base64');

    const browser = await puppeteer.launch({
      executablePath: await puppeteer.executablePath(),
      headless: true,
      args: ['--no-sandbox'],
    });
    try {
      const page = await browser.newPage();
      await page.setViewport({ width: 1280, height: 720 });
      await page.setContent(HTML);
      const textWidth = await page.evaluate(() => {
        const range = document.createRange();
        range.selectNodeContents(document.querySelector('.src'));
        return range.getBoundingClientRect().width;
      });
      // The columns of the picture that carry ink, in the picture's own pixels.
      const ink = await page.evaluate(async (src) => {
        const img = new Image();
        img.src = src;
        await img.decode();
        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const { data } = ctx.getImageData(0, 0, img.width, img.height);
        let left = img.width;
        let right = -1;
        for (let x = 0; x < img.width; x++) {
          for (let y = 0; y < img.height; y++) {
            if (data[(y * img.width + x) * 4 + 3] > 64) {
              left = Math.min(left, x);
              right = Math.max(right, x);
              break;
            }
          }
        }
        return { width: img.width, left, right };
      }, `data:image/png;base64,${png}`);

      expect(ink.left, 'ink starts where the line does').toBeLessThan(3);
      expect(Math.abs(ink.right - ink.left - textWidth) / textWidth, 'same face, same width').toBeLessThan(0.04);
    } finally {
      await browser.close();
    }
  });
});
