import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import puppeteer from 'puppeteer';
import JSZip from 'jszip';
import { Buffer } from 'node:buffer';
import { readFileSync, mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { pathToFileURL } from 'node:url';
import { exportHtmlToPptx } from '../node-exporter.js';

let browser;
let page;
let sources;
let zip;
let pptxBuffer;

beforeAll(async () => {
  browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
  page = await browser.newPage();
  await page.addScriptTag({
    type: 'module',
    content: readFileSync('src/image-processor.js', 'utf8') + '\nwindow.processImageFixture = getProcessedImage;',
  });
  await page.waitForFunction(() => typeof window.processImageFixture === 'function');
  sources = await page.evaluate(() => {
    const c = document.createElement('canvas');
    c.width = 160;
    c.height = 80;
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#ff0000';
    ctx.fillRect(0, 0, 80, 80);
    ctx.fillStyle = '#0000ff';
    ctx.fillRect(80, 0, 80, 80);
    const result = Object.fromEntries(['png', 'jpeg', 'webp'].map((type) => [type, c.toDataURL(`image/${type}`, 1)]));
    ctx.clearRect(0, 0, 20, 80);
    result.alpha = c.toDataURL('image/webp', 1);
    return result;
  });
  const cases = [
    [sources.webp, 'width:320px;height:160px', ''],
    [sources.jpeg, 'width:320px;height:160px', ''],
    [sources.png, 'width:320px;height:160px', ''],
    [sources.webp, 'width:160px;height:160px;object-fit:cover;object-position:right center;border-radius:24px', ''],
    [sources.alpha, 'width:320px;height:160px;opacity:.5', ''],
    [sources.webp, 'width:160px;height:160px;object-fit:contain', ''],
    [sources.webp, 'width:320px;height:160px', 'background'],
    [sources.webp, 'width:160px;height:160px;object-fit:none;object-position:right bottom', ''],
  ];
  const html = `<style>body{margin:0}.slide{width:400px;height:240px;position:relative;background:white}img{display:block}</style>${cases.map(([src, style, kind]) => `<section class="slide">${kind === 'background' ? `<div style="${style};background-image:url(${src});background-size:cover"></div>` : `<a href="https://example.com/image"><img src="${src}" style="${style}"></a>`}</section>`).join('')}`;
  pptxBuffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: { width: 4.1666667, height: 2.5, autoEmbedFonts: false },
  });
  zip = await JSZip.loadAsync(pptxBuffer);
}, 60000);

afterAll(async () => {
  await browser?.close();
});

async function media(slide) {
  const xml = new DOMParser().parseFromString(
    await zip.file(`ppt/slides/_rels/slide${slide}.xml.rels`).async('string'),
    'text/xml'
  );
  const rel = [...xml.getElementsByTagName('Relationship')].find((node) =>
    node.getAttribute('Type').endsWith('/image')
  );
  expect(rel).toBeDefined();
  const path = rel.getAttribute('Target').replace('../', 'ppt/');
  const data = await zip.file(path).async('nodebuffer');
  return { path, data };
}

async function pixels(slide) {
  const { data, path } = await media(slide);
  const ext = path.split('.').pop();
  return page.evaluate(
    async (src) => {
      const img = new Image();
      img.src = src;
      await img.decode();
      const c = document.createElement('canvas');
      c.width = img.width;
      c.height = img.height;
      const ctx = c.getContext('2d');
      ctx.drawImage(img, 0, 0);
      return {
        width: c.width,
        height: c.height,
        corner: [...ctx.getImageData(0, 0, 1, 1).data],
        center: [...ctx.getImageData(c.width / 2, c.height / 2, 1, 1).data],
        interior: [...ctx.getImageData(c.width * 0.75, c.height * 0.75, 1, 1).data],
        top: [...ctx.getImageData(c.width / 2, 1, 1, 1).data],
      };
    },
    `data:image/${ext};base64,${data.toString('base64')}`
  );
}

describe('image codec fidelity through HTML and PPTX', () => {
  it.each([
    ['webp', 1],
    ['jpeg', 2],
    ['png', 3],
  ])('preserves original %s bytes without upscaling', async (type, slide) => {
    const { data } = await media(slide);
    expect(data.equals(Buffer.from(sources[type].split(',')[1], 'base64'))).toBe(true);
    expect((await pixels(slide)).width).toBe(160);
  });
  it('declares real WebP media in the package content types', async () => {
    expect((await media(1)).path).toMatch(/\.webp$/);
    expect(await zip.file('[Content_Types].xml').async('string')).toContain('ContentType="image/webp"');
  });
  it('keeps a right-aligned cover crop and rounded alpha without source upscaling', async () => {
    expect((await media(4)).path).toMatch(/\.webp$/);
    const image = await pixels(4);
    expect([image.width, image.height]).toEqual([80, 80]);
    expect(image.corner[3]).toBe(0);
    expect(image.center[0]).toBeLessThan(5);
    expect(image.center[2]).toBeGreaterThan(250);
  });
  it('preserves source alpha, slide opacity, and hyperlinks', async () => {
    expect((await media(5)).data.equals(Buffer.from(sources.alpha.split(',')[1], 'base64'))).toBe(true);
    expect((await pixels(5)).corner[3]).toBe(0);
    expect(await zip.file('ppt/slides/slide5.xml').async('string')).toContain('alphaModFix amt="50000"');
    expect(await zip.file('ppt/slides/_rels/slide5.xml.rels').async('string')).toContain('https://example.com/image');
  });
  it('retains contain letterboxing as transparent padding', async () => {
    const image = await pixels(6);
    expect(image.top[3]).toBe(0);
    expect(image.center[3]).toBe(255);
    expect([image.width, image.height]).toEqual([160, 160]);
  });
  it('uses the same preservation path for CSS backgrounds', async () => {
    expect((await media(7)).data.equals(Buffer.from(sources.webp.split(',')[1], 'base64'))).toBe(true);
  });
  it('retains pixel-based object-position with object-fit none', async () => {
    const image = await pixels(8);
    expect(image.top[3]).toBe(0);
    expect(image.interior[2]).toBeGreaterThan(250);
  });
  it('preserves blob-source bytes using their MIME rather than a URL extension', async () => {
    const result = await page.evaluate(async (webp) => {
      const blob = await (await fetch(webp)).blob();
      const url = URL.createObjectURL(blob);
      try {
        return await window.processImageFixture(url, 320, 160, 0);
      } finally {
        URL.revokeObjectURL(url);
      }
    }, sources.webp);
    expect(result).toBe(sources.webp);
  });

  it('keeps a PNG MIME when the browser falls back from WebP encoding', async () => {
    const result = await page.evaluate(async (src) => {
      const original = HTMLCanvasElement.prototype.toDataURL;
      HTMLCanvasElement.prototype.toDataURL = function () {
        return original.call(this, 'image/png');
      };
      try {
        return await window.processImageFixture(src, 80, 80, 10, 'cover');
      } finally {
        HTMLCanvasElement.prototype.toDataURL = original;
      }
    }, sources.webp);
    expect(result).toMatch(/^data:image\/png;base64,/);
  });

  it('retains the readable-image fallback when fetch is unavailable', async () => {
    const result = await page.evaluate(async (src) => {
      const url = URL.createObjectURL(await (await fetch(src)).blob());
      const original = window.fetch;
      window.fetch = async () => {
        throw new TypeError('fetch is unavailable');
      };
      try {
        return await window.processImageFixture(url, 80, 80, 10, 'cover');
      } finally {
        window.fetch = original;
        URL.revokeObjectURL(url);
      }
    }, sources.webp);
    expect(result).toMatch(/^data:image\/png;base64,/);
  });

  it('keeps the 2x ceiling when downscaling a high-density raster', async () => {
    const dimensions = await page.evaluate(async (src) => {
      const data = await window.processImageFixture(src, 40, 20, 4);
      const img = new Image();
      img.src = data;
      await img.decode();
      return [img.width, img.height];
    }, sources.webp);
    expect(dimensions).toEqual([80, 40]);
  });

  it('bounds oversized originals even when no crop or mask is needed', async () => {
    const dimensions = await page.evaluate(async (src) => {
      const data = await window.processImageFixture(src, 40, 20, 0);
      const img = new Image();
      img.src = data;
      await img.decode();
      return [img.width, img.height];
    }, sources.webp);
    expect(dimensions).toEqual([80, 40]);
  });

  it('rejects empty target geometry without a hanging image job', async () => {
    expect(await page.evaluate((src) => window.processImageFixture(src, 0, 80, 0), sources.webp)).toBeNull();
  });

  it.skipIf(process.env.DOM_TO_PPTX_OFFICE_ROUNDTRIP !== '1')(
    'renders exported WebP crop and opacity in LibreOffice',
    async () => {
      const dir = mkdtempSync(path.join(tmpdir(), 'dom-to-pptx-image-codecs-'));
      try {
        writeFileSync(path.join(dir, 'images.pptx'), pptxBuffer);
        execFileSync(
          'soffice',
          [
            `-env:UserInstallation=${pathToFileURL(path.join(dir, 'profile')).href}`,
            '--headless',
            '--convert-to',
            'pdf',
            '--outdir',
            dir,
            path.join(dir, 'images.pptx'),
          ],
          { stdio: 'pipe' }
        );
        const sample = async (slide, points) => {
          const prefix = path.join(dir, `slide-${slide}`);
          execFileSync(
            'pdftoppm',
            [
              '-f',
              String(slide),
              '-l',
              String(slide),
              '-singlefile',
              '-r',
              '96',
              '-png',
              path.join(dir, 'images.pdf'),
              prefix,
            ],
            { stdio: 'pipe' }
          );
          return page.evaluate(
            async ({ src, points }) => {
              const img = new Image();
              img.src = src;
              await img.decode();
              const canvas = document.createElement('canvas');
              canvas.width = img.width;
              canvas.height = img.height;
              const ctx = canvas.getContext('2d');
              ctx.drawImage(img, 0, 0);
              return points.map(([x, y]) => [...ctx.getImageData(x, y, 1, 1).data]);
            },
            { src: 'data:image/png;base64,' + readFileSync(prefix + '.png').toString('base64'), points }
          );
        };
        const crop = await sample(4, [
          [1, 1],
          [80, 80],
          [200, 80],
        ]);
        expect(crop[0]).toEqual([255, 255, 255, 255]);
        expect(crop[1][0]).toBeLessThan(10);
        expect(crop[1][2]).toBeGreaterThan(245);
        expect(crop[2]).toEqual([255, 255, 255, 255]);
        const opacity = await sample(5, [
          [5, 40],
          [60, 40],
        ]);
        expect(opacity[0]).toEqual([255, 255, 255, 255]);
        expect(opacity[1][0]).toBeGreaterThan(245);
        expect(opacity[1][1]).toBeGreaterThan(120);
        expect(opacity[1][1]).toBeLessThan(135);
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
    },
    60000
  );
});
