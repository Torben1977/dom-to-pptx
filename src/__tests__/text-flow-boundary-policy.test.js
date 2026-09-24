import { describe, expect, it } from 'vitest';
import { exportHtmlToPptx } from '../node-exporter.js';

// Two things PowerPoint cannot hold as text at all. The converter has to say so
// per object, so the export can hand that one object over as a picture instead
// of failing — or instead of quietly producing something the reader sees as
// broken.

const html = `
  <!doctype html>
  <html>
    <head>
      <style>
        * { box-sizing: border-box; }
        body { margin: 0; font: 16px/22px Arial, sans-serif; }
        .slide { position: relative; width: 960px; height: 540px; overflow: hidden; background: #fff; }
        .prose { width: 520px; margin: 24px; }
        .prose .mark { float: left; width: 28px; font-weight: 700; }
        .columns { width: 520px; margin: 24px; }
        .columns .column { float: left; width: 240px; }
        .flagged { position: relative; padding-left: 28px; }
        .flagged .badge { position: absolute; left: 0; top: 0; font-weight: 700; }
        table { width: 520px; margin: 24px; border-collapse: collapse; }
        td { padding: 8px 10px; border-bottom: 1px solid #d1d5db; }
      </style>
    </head>
    <body>
      <section class="slide" id="float-slide">
        <p class="prose"><span class="mark">*</span>Ein Absatz, dessen Text neben dem Marker beginnt und darunter
        über die volle Breite weiterläuft, weil der Marker nur die erste Zeile verkürzt.</p>
      </section>
      <section class="slide" id="nested-float-slide">
        <p class="prose"><strong><span class="mark">*</span></strong><span>Derselbe Absatz, nur sitzt der Marker eine
        Ebene tiefer und der Text in einem Span — beides verschiebt den Fluss genauso.</span></p>
      </section>
      <section class="slide" id="table-marker-slide">
        <table><tbody><tr>
          <td class="flagged"><span class="badge">!</span>Zelle mit einem eigenen Marker</td>
          <td>offen</td>
        </tr></tbody></table>
      </section>
      <section class="slide" id="mappable-slide">
        <div class="columns"><div class="column">Linke Spalte</div><div class="column">Rechte Spalte</div></div>
        <table><tbody><tr><td><span style="transform:translateZ(0)">Malhinweis</span></td><td>offen</td></tr></tbody></table>
        <p class="flagged"><span class="badge">!</span>Ein Absatz, dessen Marker den Text nicht verschiebt.</p>
        <table><tbody><tr><td><strong>Zelle</strong> mit Auszeichnung</td><td>offen</td></tr></tbody></table>
      </section>
    </body>
  </html>
`;

const pptxOptions = { width: 10, height: 5.625, autoEmbedFonts: false };

// One browser launch per export, so every case asks its slide once and checks
// both halves at once: what the converter refuses to map, and what it tells the
// caller about it. Rasterizing keeps the deck readable but costs editable text,
// so a silent replacement would be as bad as a broken one.
async function rasterizedObjects(selector) {
  const reported = [];
  const buffer = await exportHtmlToPptx(html, {
    selector,
    pptxOptions: {
      ...pptxOptions,
      boundaryPolicy: 'rasterize',
      onBoundaryFindings: (findings) => reported.push(...findings),
    },
  });
  expect(buffer?.length, 'the export produced a deck').toBeGreaterThan(0);
  return reported.map((finding) => [finding.type, finding.container, finding.descendant]);
}

describe('text flow boundary policy', () => {
  it('replaces the block around a float inside a text flow', async () => {
    expect(await rasterizedObjects('#float-slide')).toEqual([['float-in-text-flow', 'p.prose', 'span.mark']]);
  });

  // A float displaces the inline text of the block it is laid out in, however
  // deeply either of them is nested. Asking blocks for a floated child with
  // direct text missed both halves of this — ordinary authored markup.
  it('replaces the block even when the float and the text are nested', async () => {
    expect(await rasterizedObjects('#nested-float-slide')).toEqual([['float-in-text-flow', 'p.prose', 'span.mark']]);
  });

  it('replaces the whole table when a cell holds content that needs a box of its own', async () => {
    expect(await rasterizedObjects('#table-marker-slide')).toEqual([['table-cell-needs-shape', 'table', 'span.badge']]);
  });

  // Floats that only place blocks beside each other carry no text across the
  // float, and an absolutely positioned marker is painted beside the text
  // without moving it — both map to shapes of their own and must stay text. The
  // table carries inline markup in a cell, which a native cell holds perfectly
  // well: only content that needs a box of its own is a finding. The second
  // table carries a painting hint that resolves to the identity matrix and moves
  // nothing — reading that as out-of-flow cost a whole table its editable cells.
  it('leaves floats used as columns, an absolute marker and an ordinary table alone', async () => {
    expect(await rasterizedObjects('#mappable-slide')).toEqual([]);
  });
});
