import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

const html = `<!doctype html>
<html>
  <head>
    <style>
      * { box-sizing: border-box; }
      body { margin: 0; }
      .slide { position: relative; width: 960px; height: 540px; overflow: hidden; background: #fff; }
      .box { position: absolute; left: 100px; top: 100px; width: 240px; height: 160px; }
      .wrapper { position: absolute; inset: 0; }
    </style>
  </head>
  <body>
    <section class="slide">
      <div class="box" style="position:static;margin:100px 0 0 100px;z-index:999;background:#d92d20"></div>
      <div class="box" style="background:#2563eb"></div>
    </section>
    <section class="slide">
      <div class="wrapper"><div class="box" style="z-index:10;background:#d92d20"></div></div>
      <div class="wrapper"><div class="box" style="z-index:1;background:#2563eb"></div></div>
    </section>
    <section class="slide">
      <div class="wrapper" style="z-index:0"><div class="box" style="z-index:10;background:#d92d20"></div></div>
      <div class="box" style="z-index:1;background:#2563eb"></div>
    </section>
    <section class="slide">
      <div class="wrapper" style="opacity:.99"><div class="box" style="z-index:10;background:#d92d20"></div></div>
      <div class="box" style="z-index:1;background:#2563eb"></div>
    </section>
    <section class="slide">
      <div class="wrapper" style="transform:translateX(0)"><div class="box" style="z-index:10;background:#d92d20"></div></div>
      <div class="box" style="z-index:1;background:#2563eb"></div>
    </section>
  </body>
</html>`;

let zip;

beforeAll(async () => {
  const buffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: { width: 10, height: 5.625, autoEmbedFonts: false },
  });
  zip = await JSZip.loadAsync(buffer);
}, 40_000);

async function fillOrder(slideNumber) {
  const xml = await zip.file(`ppt/slides/slide${slideNumber}.xml`).async('string');
  const documentNode = new DOMParser().parseFromString(xml, 'text/xml');
  const shapeTree = documentNode.getElementsByTagName('p:spTree')[0];
  return Array.from(shapeTree.children)
    .filter((node) => node.tagName === 'p:sp')
    .flatMap((shape) =>
      Array.from(shape.getElementsByTagName('a:solidFill')).flatMap((fill) =>
        Array.from(fill.getElementsByTagName('a:srgbClr')).map((color) => color.getAttribute('val'))
      )
    )
    .filter((color) => color === 'D92D20' || color === '2563EB');
}

describe('CSS stacking contexts', () => {
  it('ignores z-index on a non-positioned block', async () => {
    expect(await fillOrder(1)).toEqual(['D92D20', '2563EB']);
  });

  it('lets a positioned child escape a wrapper that is not a stacking context', async () => {
    expect(await fillOrder(2)).toEqual(['2563EB', 'D92D20']);
  });

  it('keeps a positioned child inside a real parent stacking context', async () => {
    expect(await fillOrder(3)).toEqual(['D92D20', '2563EB']);
  });

  it('keeps a positioned child inside an opacity stacking context', async () => {
    expect(await fillOrder(4)).toEqual(['D92D20', '2563EB']);
  });

  it('keeps a positioned child inside a transform stacking context', async () => {
    expect(await fillOrder(5)).toEqual(['D92D20', '2563EB']);
  });
});
