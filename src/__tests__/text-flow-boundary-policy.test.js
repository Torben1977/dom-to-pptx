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
      <section class="slide" id="table-marker-slide">
        <table><tbody><tr>
          <td class="flagged"><span class="badge">!</span>Zelle mit einem eigenen Marker</td>
          <td>offen</td>
        </tr></tbody></table>
      </section>
      <section class="slide" id="mappable-slide">
        <div class="columns"><div class="column">Linke Spalte</div><div class="column">Rechte Spalte</div></div>
        <p class="flagged"><span class="badge">!</span>Ein Absatz, dessen Marker den Text nicht verschiebt.</p>
        <table><tbody><tr><td><strong>Zelle</strong> mit Auszeichnung</td><td>offen</td></tr></tbody></table>
      </section>
    </body>
  </html>
`;

const pptxOptions = { width: 10, height: 5.625, autoEmbedFonts: false };

const exportSlide = (selector, boundaryPolicy) =>
  exportHtmlToPptx(html, { selector, pptxOptions: { ...pptxOptions, boundaryPolicy } });

describe('text flow boundary policy', () => {
  it('reports a float inside a text flow, naming the block and the float', async () => {
    await expect(exportSlide('#float-slide', 'error')).rejects.toThrow(
      /DOM_TO_PPTX_UNSUPPORTED_BOUNDARY.*float-in-text-flow.*p\.prose.*span\.mark/s
    );
  });

  it('reports out-of-flow content in a table cell against the table, which is what can be replaced', async () => {
    await expect(exportSlide('#table-marker-slide', 'error')).rejects.toThrow(
      /DOM_TO_PPTX_UNSUPPORTED_BOUNDARY.*table-cell-needs-shape.*table.*span\.badge/s
    );
  });

  // Floats that only place blocks beside each other carry no text across the
  // float, and an absolutely positioned marker is painted beside the text
  // without moving it — both map to shapes of their own and must stay text. The
  // table carries inline markup in a cell, which a native cell holds perfectly
  // well: only content that needs a box of its own is a finding.
  it('leaves floats used as columns, an absolute marker and an ordinary table alone', async () => {
    await expect(exportSlide('#mappable-slide', 'error')).resolves.toBeInstanceOf(Buffer);
  });
});
