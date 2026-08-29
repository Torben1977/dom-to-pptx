import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

const EMU_PER_PX = 9_525;

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
        .paint-wrapper {
          position: absolute;
          inset: 0;
          z-index: auto;
        }
        .paint-box {
          position: absolute;
          left: 100px;
          top: 100px;
          width: 240px;
          height: 160px;
        }
        .clip-parent {
          position: absolute;
          left: 100px;
          top: 100px;
          width: 200px;
          height: 120px;
          overflow: hidden;
        }
        .clip-child {
          position: absolute;
          left: 150px;
          top: 20px;
          width: 120px;
          height: 60px;
          background: #d92d20;
        }
        .clip-text-child {
          position: absolute;
          left: 150px;
          top: 80px;
          width: 120px;
          height: 30px;
          background: #f79009;
          color: #ffffff;
        }
        .transform-parent {
          position: absolute;
          left: 300px;
          top: 180px;
          width: 200px;
          height: 200px;
          transform: rotate(20deg);
          transform-origin: center;
        }
        .transform-child {
          position: absolute;
          left: 50px;
          top: 50px;
          width: 100px;
          height: 100px;
          transform: rotate(10deg);
          transform-origin: center;
          background: #6f42c1;
        }
        table {
          position: absolute;
          left: 80px;
          top: 80px;
          width: 600px;
          border-collapse: collapse;
          table-layout: auto;
          font: 20px/1.2 Arial, sans-serif;
        }
        td { padding: 0; border: 0; }
        .asymmetric-card {
          position: absolute;
          left: 120px;
          top: 100px;
          width: 360px;
          height: 180px;
          padding: 24px;
          background: #eef4ee;
          border-radius: 48px 0 0 0;
          font: 700 28px/1.2 Arial, sans-serif;
        }
        .asymmetric-border-card {
          position: absolute;
          left: 120px;
          top: 100px;
          width: 360px;
          height: 180px;
          padding: 24px;
          background: transparent;
          border: 6px solid #b91c1c;
          border-radius: 48px 0 0 0;
          font: 700 28px/1.2 Arial, sans-serif;
        }
        x-status-icon {
          position: absolute;
          left: 120px;
          top: 100px;
          display: block;
          width: 120px;
          height: 120px;
        }
        x-decorated-icon,
        x-mixed-icon {
          position: absolute;
          left: 120px;
          top: 100px;
          display: block;
          width: 120px;
          height: 120px;
        }
        x-decorated-icon {
          padding: 20px;
          background: #fee2e2;
          border: 4px solid #b91c1c;
        }
      </style>
    </head>
    <body>
      <section class="slide">
        <div class="paint-wrapper">
          <div class="paint-box" style="z-index:10;background:#d92d20"></div>
        </div>
        <div class="paint-wrapper">
          <div class="paint-box" style="z-index:1;background:#2563eb"></div>
        </div>
      </section>

      <section class="slide">
        <div class="clip-parent">
          <div class="clip-child"></div>
          <div class="clip-text-child">CLIPPED_TEXT</div>
        </div>
      </section>

      <section class="slide">
        <div class="transform-parent">
          <div class="transform-child"></div>
        </div>
      </section>

      <section class="slide">
        <table>
          <tbody>
            <tr>
              <td rowspan="2" style="width:100px">A</td>
              <td colspan="2" style="width:500px">B</td>
            </tr>
            <tr>
              <td style="width:200px">C</td>
              <td style="width:300px">D</td>
            </tr>
          </tbody>
        </table>
      </section>

      <section class="slide">
        <div class="asymmetric-card">ASYMMETRIC_CARD</div>
      </section>

      <section class="slide">
        <x-status-icon></x-status-icon>
      </section>

      <section class="slide">
        <table style="table-layout:fixed">
          <colgroup><col style="width:100px"></colgroup>
          <tbody><tr><td>A</td><td>B</td><td>C</td></tr></tbody>
        </table>
      </section>

      <section class="slide">
        <table>
          <tbody>
            <tr style="visibility:collapse"><td>SECRET_COLLAPSED_ROW</td></tr>
            <tr><td>VISIBLE_ROW</td></tr>
          </tbody>
        </table>
      </section>

      <section class="slide">
        <table>
          <tbody style="display:none"><tr><td>SECRET_HIDDEN_GROUP</td></tr></tbody>
          <tbody><tr><td>VISIBLE_GROUP_ROW</td></tr></tbody>
        </table>
      </section>

      <section class="slide">
        <x-decorated-icon></x-decorated-icon>
      </section>

      <section class="slide">
        <x-mixed-icon></x-mixed-icon>
      </section>

      <section class="slide">
        <div class="asymmetric-border-card">ASYMMETRIC_BORDER_CARD</div>
      </section>

      <script>
        customElements.define(
          'x-status-icon',
          class extends HTMLElement {
            connectedCallback() {
              const shadow = this.attachShadow({ mode: 'open' });
              shadow.innerHTML =
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120">' +
                '<circle cx="60" cy="60" r="54" fill="#0f766e" />' +
                '<path d="M32 62l18 18 38-42" fill="none" stroke="#fff" stroke-width="12" />' +
                '</svg>';
            }
          }
        );
        customElements.define(
          'x-decorated-icon',
          class extends HTMLElement {
            connectedCallback() {
              const shadow = this.attachShadow({ mode: 'open' });
              shadow.innerHTML =
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 72 72">' +
                '<circle cx="36" cy="36" r="32" fill="#0f766e" />' +
                '</svg>';
            }
          }
        );
        customElements.define(
          'x-mixed-icon',
          class extends HTMLElement {
            connectedCallback() {
              const shadow = this.attachShadow({ mode: 'open' });
              shadow.innerHTML =
                'VISIBLE_SHADOW_TEXT' +
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120">' +
                '<circle cx="60" cy="60" r="54" fill="#0f766e" />' +
                '</svg>';
            }
          }
        );
      </script>
    </body>
  </html>
`;

let zip;

beforeAll(async () => {
  const buffer = await exportHtmlToPptx(html, {
    selector: '.slide',
    pptxOptions: {
      width: 10,
      height: 5.625,
      autoEmbedFonts: false,
    },
  });
  zip = await JSZip.loadAsync(buffer);
}, 40_000);

async function slideDocument(slideNumber) {
  const xml = await zip.file(`ppt/slides/slide${slideNumber}.xml`).async('string');
  return new DOMParser().parseFromString(xml, 'text/xml');
}

async function slideRelationshipsDocument(slideNumber) {
  const xml = await zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`).async('string');
  return new DOMParser().parseFromString(xml, 'text/xml');
}

function drawingShapes(documentNode) {
  const shapeTree = documentNode.getElementsByTagName('p:spTree')[0];
  return Array.from(shapeTree.children).filter((node) => node.tagName === 'p:sp');
}

function shapeWithFill(documentNode, color) {
  return drawingShapes(documentNode).find((shape) =>
    Array.from(shape.getElementsByTagName('a:solidFill')).some((fill) =>
      Array.from(fill.getElementsByTagName('a:srgbClr')).some((node) => node.getAttribute('val') === color)
    )
  );
}

function transformOf(shape) {
  const xfrm = shape.getElementsByTagName('a:xfrm')[0];
  const off = xfrm.getElementsByTagName('a:off')[0];
  const ext = xfrm.getElementsByTagName('a:ext')[0];
  return {
    x: Number(off.getAttribute('x')),
    y: Number(off.getAttribute('y')),
    w: Number(ext.getAttribute('cx')),
    h: Number(ext.getAttribute('cy')),
    rotation: Number(xfrm.getAttribute('rot') || 0),
  };
}

describe('known renderer boundaries', () => {
  it('keeps both overlapping shapes before evaluating their unresolved CSS paint order', async () => {
    const documentNode = await slideDocument(1);
    expect(shapeWithFill(documentNode, 'D92D20')).toBeDefined();
    expect(shapeWithFill(documentNode, '2563EB')).toBeDefined();
  });

  it('preserves CSS paint order across wrappers that do not create stacking contexts', async () => {
    const documentNode = await slideDocument(1);
    const shapes = drawingShapes(documentNode);
    const red = shapeWithFill(documentNode, 'D92D20');
    const blue = shapeWithFill(documentNode, '2563EB');

    expect(red).toBeDefined();
    expect(blue).toBeDefined();
    expect(shapes.indexOf(red)).toBeGreaterThan(shapes.indexOf(blue));
  });

  it('keeps the overflowing child before evaluating its unresolved ancestor clipping', async () => {
    const documentNode = await slideDocument(2);
    expect(shapeWithFill(documentNode, 'D92D20')).toBeDefined();
  });

  it('clips an axis-aligned empty solid child to its overflow-hidden ancestor', async () => {
    const documentNode = await slideDocument(2);
    const child = shapeWithFill(documentNode, 'D92D20');

    expect(child).toBeDefined();
    expect(transformOf(child)).toMatchObject({
      x: 250 * EMU_PER_PX,
      y: 120 * EMU_PER_PX,
      w: 50 * EMU_PER_PX,
      h: 60 * EMU_PER_PX,
    });
  });

  it.fails('clips text-bearing children to the geometry of their overflow-hidden ancestor', async () => {
    const documentNode = await slideDocument(2);
    const child = shapeWithFill(documentNode, 'F79009');

    expect(child).toBeDefined();
    expect(transformOf(child)).toMatchObject({
      x: 250 * EMU_PER_PX,
      y: 180 * EMU_PER_PX,
      w: 50 * EMU_PER_PX,
      h: 30 * EMU_PER_PX,
    });
  });

  it('keeps the transformed child before evaluating its unresolved cumulative transform', async () => {
    const documentNode = await slideDocument(3);
    const child = shapeWithFill(documentNode, '6F42C1');
    expect(child).toBeDefined();
    const transform = transformOf(child);
    expect(Number.isFinite(transform.x)).toBe(true);
    expect(Number.isFinite(transform.y)).toBe(true);
    expect(transform.w).toBeGreaterThan(0);
    expect(transform.h).toBeGreaterThan(0);
  });

  it.fails('maps cumulative ancestor and child rotations to the editable PowerPoint shape', async () => {
    const documentNode = await slideDocument(3);
    const child = shapeWithFill(documentNode, '6F42C1');

    expect(child).toBeDefined();
    expect(transformOf(child)).toMatchObject({
      x: 350 * EMU_PER_PX,
      y: 230 * EMU_PER_PX,
      w: 100 * EMU_PER_PX,
      h: 100 * EMU_PER_PX,
      rotation: 30 * 60_000,
    });
  });

  it('keeps asymmetric-card text before evaluating its unresolved corner geometry', async () => {
    const documentNode = await slideDocument(5);
    expect(
      drawingShapes(documentNode).some((shape) =>
        Array.from(shape.getElementsByTagName('a:t')).some((node) => node.textContent === 'ASYMMETRIC_CARD')
      )
    ).toBe(true);
  });

  it('derives the complete column grid when rowspans occupy columns in later rows', async () => {
    const documentNode = await slideDocument(4);
    const gridWidths = Array.from(documentNode.getElementsByTagName('a:gridCol'), (column) =>
      Number(column.getAttribute('w'))
    );

    expect(gridWidths).toEqual([100 * EMU_PER_PX, 200 * EMU_PER_PX, 300 * EMU_PER_PX]);
  });

  it('preserves asymmetric corner radii when the same element also owns editable text', async () => {
    const documentNode = await slideDocument(5);
    const textShape = drawingShapes(documentNode).find((shape) =>
      Array.from(shape.getElementsByTagName('a:t')).some((node) => node.textContent === 'ASYMMETRIC_CARD')
    );
    const geometry = textShape?.getElementsByTagName('a:prstGeom')[0]?.getAttribute('prst');
    const textUsesCustomGeometry = textShape?.getElementsByTagName('a:custGeom').length > 0;
    const separateVisualLayer =
      documentNode.getElementsByTagName('p:pic').length > 0 ||
      drawingShapes(documentNode).some(
        (shape) => shape !== textShape && shape.getElementsByTagName('a:custGeom').length > 0
      );

    // Either one custom editable geometry owns the text, or editable rectangular
    // text is overlaid on a separate visual layer. A roundRect alone is wrong
    // because it rounds all four corners.
    expect(textUsesCustomGeometry || (geometry === 'rect' && separateVisualLayer)).toBe(true);
  });

  it('captures an SVG-only open Shadow-DOM icon as a non-empty vector picture at the browser size', async () => {
    const documentNode = await slideDocument(6);
    const relationships = await slideRelationshipsDocument(6);
    const pictureEvidence = await Promise.all(
      Array.from(documentNode.getElementsByTagName('p:pic'), async (picture) => {
        const transform = picture.getElementsByTagName('a:xfrm')[0];
        const ext = transform?.getElementsByTagName('a:ext')[0];
        const embedId = picture.getElementsByTagName('asvg:svgBlip')[0]?.getAttribute('r:embed');
        const relationship = Array.from(relationships.getElementsByTagName('Relationship')).find(
          (entry) => entry.getAttribute('Id') === embedId
        );
        const target = relationship?.getAttribute('Target');
        const mediaPath = target ? `ppt/${target.replace(/^\.\.\//, '')}` : null;
        const media = mediaPath && zip.file(mediaPath) ? await zip.file(mediaPath).async('string') : null;

        return {
          width: Number(ext?.getAttribute('cx')),
          height: Number(ext?.getAttribute('cy')),
          mediaPath,
          hasSubstantialMedia: (media?.length || 0) > 256,
          hasVectorElements: /<(?:circle|path)\b/i.test(media || ''),
          hasGreenFill: /(?:#0f766e|rgb\(15,\s*118,\s*110\))/i.test(media || ''),
        };
      })
    );

    expect(pictureEvidence).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          width: 120 * EMU_PER_PX,
          height: 120 * EMU_PER_PX,
          mediaPath: expect.stringMatching(/\.svg$/),
          hasSubstantialMedia: true,
          hasVectorElements: true,
          hasGreenFill: true,
        }),
      ])
    );
  });

  it('preserves a decorated Shadow-DOM host and places its SVG in the rendered content box', async () => {
    const documentNode = await slideDocument(10);
    const hostShape = shapeWithFill(documentNode, 'FEE2E2');
    expect(hostShape).toBeDefined();
    expect(transformOf(hostShape)).toMatchObject({
      x: 120 * EMU_PER_PX,
      y: 100 * EMU_PER_PX,
      w: 120 * EMU_PER_PX,
      h: 120 * EMU_PER_PX,
    });

    const vectorPicture = Array.from(documentNode.getElementsByTagName('p:pic')).find(
      (picture) => picture.getElementsByTagName('asvg:svgBlip').length === 1
    );
    expect(vectorPicture).toBeDefined();
    const transform = vectorPicture.getElementsByTagName('a:xfrm')[0];
    const offset = transform.getElementsByTagName('a:off')[0];
    const extent = transform.getElementsByTagName('a:ext')[0];
    expect({
      x: Number(offset.getAttribute('x')),
      y: Number(offset.getAttribute('y')),
      w: Number(extent.getAttribute('cx')),
      h: Number(extent.getAttribute('cy')),
    }).toEqual({
      x: 144 * EMU_PER_PX,
      y: 124 * EMU_PER_PX,
      w: 72 * EMU_PER_PX,
      h: 72 * EMU_PER_PX,
    });
  });

  it('does not classify visible Shadow-DOM text plus an SVG as SVG-only', async () => {
    const documentNode = await slideDocument(11);
    expect(documentNode.getElementsByTagName('asvg:svgBlip')).toHaveLength(0);
  });

  it.fails('preserves visible mixed Shadow-DOM text instead of dropping the shadow tree', async () => {
    const documentNode = await slideDocument(11);
    expect(documentNode.documentElement.textContent).toContain('VISIBLE_SHADOW_TEXT');
  });

  it('preserves a transparent border-only asymmetric card behind editable text', async () => {
    const documentNode = await slideDocument(12);
    const relationships = await slideRelationshipsDocument(12);
    const textShape = drawingShapes(documentNode).find((shape) =>
      Array.from(shape.getElementsByTagName('a:t')).some((node) => node.textContent === 'ASYMMETRIC_BORDER_CARD')
    );
    expect(textShape).toBeDefined();
    expect(textShape.getElementsByTagName('a:prstGeom')[0]?.getAttribute('prst')).toBe('rect');

    const vectorPicture = Array.from(documentNode.getElementsByTagName('p:pic')).find(
      (picture) => picture.getElementsByTagName('asvg:svgBlip').length === 1
    );
    expect(vectorPicture).toBeDefined();
    const embedId = vectorPicture.getElementsByTagName('asvg:svgBlip')[0].getAttribute('r:embed');
    const relationship = Array.from(relationships.getElementsByTagName('Relationship')).find(
      (entry) => entry.getAttribute('Id') === embedId
    );
    const mediaPath = `ppt/${relationship.getAttribute('Target').replace(/^\.\.\//, '')}`;
    const media = await zip.file(mediaPath).async('string');
    expect(media).toMatch(/<path[^>]+fill="none"/i);
    expect(media).toMatch(/stroke="#b91c1c"/i);
  });

  it('completes a partial colgroup with the browser-measured implicit columns', async () => {
    const documentNode = await slideDocument(7);
    const gridWidths = Array.from(documentNode.getElementsByTagName('a:gridCol'), (column) =>
      Number(column.getAttribute('w'))
    );
    expect(gridWidths).toEqual([100 * EMU_PER_PX, 250 * EMU_PER_PX, 250 * EMU_PER_PX]);
  });

  it('omits visibility-collapsed table rows from the native PowerPoint table', async () => {
    const documentNode = await slideDocument(8);
    expect(documentNode.getElementsByTagName('a:tr')).toHaveLength(1);
    expect(documentNode.documentElement.textContent).toContain('VISIBLE_ROW');
    expect(documentNode.documentElement.textContent).not.toContain('SECRET_COLLAPSED_ROW');
  });

  it('omits rows inside display-none table groups from the native PowerPoint table', async () => {
    const documentNode = await slideDocument(9);
    expect(documentNode.getElementsByTagName('a:tr')).toHaveLength(1);
    expect(documentNode.documentElement.textContent).toContain('VISIBLE_GROUP_ROW');
    expect(documentNode.documentElement.textContent).not.toContain('SECRET_HIDDEN_GROUP');
  });
});
