import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportHtmlToPptx } from '../node-exporter.js';

const html = `
  <!doctype html>
  <html>
    <head>
      <base href="https://example.com/base/">
      <style>
        body { margin: 0; }
        .slide {
          width: 960px;
          height: 540px;
          position: relative;
          overflow: hidden;
          background: #fff;
        }
        .box {
          position: absolute;
          left: 40px;
          top: 40px;
          width: 500px;
          font: 32px/1.25 Arial, sans-serif;
        }
        .semantic-card {
          display: block;
          position: absolute;
          left: 40px;
          top: 40px;
          width: 400px;
          height: 180px;
          padding: 20px;
          background: #eef2ff;
        }
        .semantic-card h2,
        .semantic-card p { margin: 0; }
        table {
          position: absolute;
          left: 40px;
          top: 40px;
          width: 600px;
          table-layout: fixed;
          border-collapse: collapse;
        }
        td, th { border: 1px solid #111; }
        img { width: 120px; height: 80px; }
        .shape-link {
          display: block;
          position: absolute;
          left: 40px;
          top: 40px;
          width: 240px;
          height: 120px;
          background: #d92d20;
          border-radius: 12px;
        }
        .asymmetric-shape-link {
          border-radius: 32px 0 0 0;
        }
      </style>
    </head>
    <body>
      <div class="slide">
        <div class="box">Visible <span style="display:none">SECRET_HIDDEN</span></div>
      </div>

      <div class="slide">
        <div class="box" style="opacity:.5">plain <b>bold</b> tail</div>
      </div>

      <div class="slide">
        <kpi-card class="semantic-card">
          <h2>Revenue</h2>
          <p>42 percent growth</p>
        </kpi-card>
      </div>

      <div class="slide">
        <table>
          <colgroup>
            <col style="width:200px">
            <col style="width:400px">
          </colgroup>
          <tr><th colspan="2">Header</th></tr>
          <tr><td>A</td><td>B</td></tr>
        </table>
      </div>

      <div class="slide">
        <div class="box"><a href="docs/report.html">Relative report</a></div>
      </div>

      <div class="slide">
        <a href="https://example.com/image">
          <img
            alt="linked image"
            src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='120' height='80'%3E%3Crect width='120' height='80' fill='red'/%3E%3C/svg%3E"
          >
        </a>
      </div>

      <div class="slide" data-pptx-notes="Root presenter note.">
        <div class="box">Root-notes slide</div>
      </div>

      <div class="slide">
        <div class="box">Attribute-notes slide</div>
        <template data-pptx-notes="Attribute presenter note."></template>
      </div>

      <div class="slide">
        <a class="shape-link" href="https://example.com/shape" aria-label="Shape link"></a>
      </div>

      <div class="slide">
        <div class="box" style="visibility:hidden">
          SECRET_INHERITED
          <span style="display:block;visibility:visible">VISIBLE_OVERRIDE</span>
        </div>
      </div>

      <div class="slide">
        <a
          class="shape-link asymmetric-shape-link"
          href="https://example.com/asymmetric-shape"
          aria-label="Asymmetric shape link"
        ></a>
      </div>

      <div class="slide">
        <div style="opacity:.5">
          <img
            alt="half transparent image"
            src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='120' height='80'%3E%3Crect width='120' height='80' fill='red'/%3E%3C/svg%3E"
          >
        </div>
      </div>

      <div class="slide">
        <div style="opacity:.5">
          <svg width="120" height="80" viewBox="0 0 120 80">
            <rect width="120" height="80" fill="#d92d20"></rect>
          </svg>
        </div>
      </div>

      <div class="slide">
        <div style="opacity:.5">
          <canvas id="opacity-canvas" width="120" height="80"></canvas>
        </div>
      </div>

      <div class="slide">
        <div style="opacity:.5">
          <div
            style="width:120px;height:80px;background-size:cover;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='120' height='80'%3E%3Crect width='120' height='80' fill='red'/%3E%3C/svg%3E&quot;)"
          ></div>
        </div>
      </div>

      <div class="slide">
        <a
          href="https://example.com/linked-text-card"
          style="position:absolute;left:120px;top:100px;display:block;width:420px;height:180px;padding:24px;background:#eef4ee;color:#0f172a;text-decoration:none"
        ><strong>Linked title</strong><span style="display:block">Linked body</span></a>
      </div>

      <script>
        const opacityCanvas = document.getElementById('opacity-canvas');
        const opacityContext = opacityCanvas.getContext('2d');
        opacityContext.fillStyle = '#d92d20';
        opacityContext.fillRect(0, 0, 120, 80);
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
}, 40000);

async function slideXml(slideNumber) {
  return zip.file(`ppt/slides/slide${slideNumber}.xml`).async('string');
}

async function slideRelationships(slideNumber) {
  return zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`).async('string');
}

async function notesXml(slideNumber) {
  return zip.file(`ppt/notesSlides/notesSlide${slideNumber}.xml`).async('string');
}

function textRun(xml, text) {
  const textIndex = xml.indexOf(`<a:t>${text}</a:t>`);
  expect(textIndex).toBeGreaterThan(-1);
  const runStart = xml.lastIndexOf('<a:r>', textIndex);
  const runEnd = xml.indexOf('</a:r>', textIndex);
  expect(runStart).toBeGreaterThan(-1);
  expect(runEnd).toBeGreaterThan(textIndex);
  return xml.slice(runStart, runEnd + '</a:r>'.length);
}

describe('adversarial DOM semantics', () => {
  it('keeps hidden inline text out of a shared text box', async () => {
    const hiddenXml = await slideXml(1);
    expect(hiddenXml).toContain('<a:t>Visible</a:t>');
    expect(hiddenXml).not.toContain('SECRET_HIDDEN');
  });

  it('compounds parent opacity into rich-text children', async () => {
    const opacityXml = await slideXml(2);
    expect(textRun(opacityXml, 'plain ')).toContain('<a:alpha val="50000"/>');
    expect(textRun(opacityXml, 'bold')).toContain('<a:alpha val="50000"/>');
    expect(textRun(opacityXml, ' tail')).toContain('<a:alpha val="50000"/>');
  });

  it('keeps semantic custom-element content editable instead of classifying every hyphenated tag as an icon', async () => {
    const xml = await slideXml(3);
    expect(xml).toContain('<a:t>Revenue</a:t>');
    expect(xml).toContain('<a:t>42 percent growth</a:t>');
  });

  it('derives unequal table columns from the rendered grid when the first row is a colspan', async () => {
    const xml = await slideXml(4);
    const gridWidths = Array.from(xml.matchAll(/<a:gridCol w="(\d+)"/g), (match) => Number(match[1]));

    expect(gridWidths).toEqual([1_905_000, 3_810_000]);
  });

  it('resolves relative text links against the document base URL', async () => {
    const textRelationships = await slideRelationships(5);
    expect(textRelationships).toContain('Target="https://example.com/base/docs/report.html"');
  });

  it('preserves hyperlinks wrapped around images', async () => {
    const imageRelationships = await slideRelationships(6);
    expect(imageRelationships).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
    );
    expect(imageRelationships).toContain('Target="https://example.com/image"');
  });

  it('reads speaker-note text from the slide root attribute', async () => {
    expect(await notesXml(7)).toContain('<a:t>Root presenter note.</a:t>');
  });

  it('reads speaker-note text from a nested data-pptx-notes attribute', async () => {
    expect(await notesXml(8)).toContain('<a:t>Attribute presenter note.</a:t>');
  });

  it('preserves hyperlinks on shape-only anchors', async () => {
    const shapeRelationships = await slideRelationships(9);
    expect(shapeRelationships).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
    );
    expect(shapeRelationships).toContain('Target="https://example.com/shape"');
  });

  it('keeps an explicitly visible descendant of a visibility-hidden parent', async () => {
    const xml = await slideXml(10);
    expect(xml).toContain('VISIBLE_OVERRIDE');
    expect(xml).not.toContain('SECRET_INHERITED');
  });

  it('emits one linked visual object for an empty anchor with asymmetric corners', async () => {
    const xml = await slideXml(11);
    const relationships = await slideRelationships(11);

    expect(xml.match(/<p:pic>/g) || []).toHaveLength(1);
    expect(xml.match(/<a:hlinkClick\b/g) || []).toHaveLength(1);
    expect(
      relationships.match(
        /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"/g
      ) || []
    ).toHaveLength(1);
    expect(relationships.match(/Target="https:\/\/example\.com\/asymmetric-shape"/g) || []).toHaveLength(1);
  });

  it.each([
    ['IMG', 12],
    ['SVG', 13],
    ['canvas', 14],
    ['background image', 15],
  ])('preserves inherited opacity on media exported from a %s element', async (_elementName, slideNumber) => {
    const xml = await slideXml(slideNumber);
    expect(xml).toContain('<a:alphaModFix amt="50000"/>');
  });

  it('emits only valid hyperlinks for every editable part of a linked text card', async () => {
    const xml = await slideXml(16);
    const relationships = await slideRelationships(16);
    const slideDocument = new DOMParser().parseFromString(xml, 'text/xml');
    const relationshipsDocument = new DOMParser().parseFromString(relationships, 'text/xml');
    const clickIds = Array.from(slideDocument.getElementsByTagName('a:hlinkClick'), (node) =>
      node.getAttribute('r:id')
    );
    const hyperlinkRelationships = Array.from(relationshipsDocument.getElementsByTagName('Relationship')).filter(
      (node) => node.getAttribute('Type')?.endsWith('/hyperlink')
    );
    const relationshipsById = new Map(hyperlinkRelationships.map((node) => [node.getAttribute('Id'), node]));

    expect(clickIds.length).toBeGreaterThan(0);
    expect(clickIds).not.toContain('rIdundefined');
    expect(hyperlinkRelationships).toHaveLength(new Set(clickIds).size);
    for (const clickId of clickIds) {
      const relationship = relationshipsById.get(clickId);
      expect(relationship).toBeDefined();
      expect(relationship.getAttribute('Target')).toBe('https://example.com/linked-text-card');
    }
  });
});
