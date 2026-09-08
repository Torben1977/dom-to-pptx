import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import path from 'node:path';
import { exportHtmlToPptx } from '../node-exporter.js';

const FIXTURE = path.resolve('src/__tests__/fixtures/flex-auto-width-regression.html');
const PROCESS_LABELS = ['Wechselwirkung', 'Ausgangswerte', 'Umsetzung', 'Review'];
const AUTO_LABELS = ['Entscheidung jetzt', ...PROCESS_LABELS];

function shapeFor(xml, text) {
  const textIndex = xml.indexOf(`<a:t>${text}</a:t>`);
  expect(textIndex, `text '${text}'`).toBeGreaterThan(-1);
  const shapeStart = xml.lastIndexOf('<p:sp>', textIndex);
  const shapeEnd = xml.indexOf('</p:sp>', textIndex);
  return xml.slice(shapeStart, shapeEnd);
}

function width(shape) {
  return Number(shape.match(/<a:ext cx="(\d+)"/)?.[1]);
}

function geometry(shape) {
  const x = Number(shape.match(/<a:off x="(\d+)"/)?.[1]);
  const cx = width(shape);
  return { x, width: cx, right: x + cx };
}

function cssPxToEmu(px) {
  return Math.round((px / 144) * 914_400);
}

describe('auto-width flex text fidelity', () => {
  it('adds a bounded reserve to painted auto-width flex labels', async () => {
    const buffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const unreservedBuffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false, _textFitReserveRatio: 0 },
    });
    const zip = await JSZip.loadAsync(buffer);
    const unreservedZip = await JSZip.loadAsync(unreservedBuffer);
    const paintedXml = await zip.file('ppt/slides/slide1.xml').async('string');
    const unreservedXml = await unreservedZip.file('ppt/slides/slide1.xml').async('string');

    for (const text of AUTO_LABELS) {
      const painted = shapeFor(paintedXml, text);
      const unreserved = width(shapeFor(unreservedXml, text));
      expect(width(painted), text).toBeGreaterThan(unreserved * 1.055);
      expect(width(painted), text).toBeLessThan(unreserved * 1.121);
      expect(painted).toContain('wrap="square"');
      expect(painted).not.toContain('<a:normAutofit');
    }

    for (let index = 0; index < PROCESS_LABELS.length - 1; index++) {
      const current = geometry(shapeFor(paintedXml, PROCESS_LABELS[index]));
      const next = geometry(shapeFor(paintedXml, PROCESS_LABELS[index + 1]));
      expect(current.right, `${PROCESS_LABELS[index]} overlaps ${PROCESS_LABELS[index + 1]}`).toBeLessThanOrEqual(
        next.x
      );
    }
    expect(width(shapeFor(paintedXml, 'Review'))).toBeGreaterThan(width(shapeFor(unreservedXml, 'Review')) * 1.115);
  }, 40_000);

  it('reserves enough width for several bold auto-width pills in one row', async () => {
    const buffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const unreservedBuffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false, _textFitReserveRatio: 0 },
    });
    const zip = await JSZip.loadAsync(buffer);
    const unreservedZip = await JSZip.loadAsync(unreservedBuffer);
    const xml = await zip.file('ppt/slides/slide4.xml').async('string');
    const unreservedXml = await unreservedZip.file('ppt/slides/slide4.xml').async('string');

    for (const text of AUTO_LABELS) {
      const pill = shapeFor(xml, text);
      const unreserved = width(shapeFor(unreservedXml, text));
      expect(width(pill), text).toBeGreaterThan(unreserved * 1.055);
      expect(width(pill), text).toBeLessThan(unreserved * 1.121);
      expect(pill).toContain('wrap="square"');
      expect(pill).not.toContain('<a:normAutofit');
    }

    for (let index = 0; index < AUTO_LABELS.length - 1; index++) {
      const current = geometry(shapeFor(xml, AUTO_LABELS[index]));
      const next = geometry(shapeFor(xml, AUTO_LABELS[index + 1]));
      expect(current.right, `${AUTO_LABELS[index]} overlaps ${AUTO_LABELS[index + 1]}`).toBeLessThanOrEqual(next.x + 1);
    }
    expect(width(shapeFor(xml, 'Review'))).toBeGreaterThan(width(shapeFor(unreservedXml, 'Review')) * 1.115);
  }, 40_000);

  it('preserves constrained, explicit-break, and ordinary paragraph wrapping', async () => {
    const buffer = await exportHtmlToPptx(FIXTURE, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide3.xml').async('string');

    const narrow = shapeFor(xml, 'Dieser Text soll weiterhin regulär zwischen Wörtern umbrechen können.');
    const normal = shapeFor(
      xml,
      'Normaler Absatztext bleibt in seiner begrenzten Textbox und verhält sich unverändert.'
    );
    const explicit = shapeFor(xml, 'Erste Zeile');

    expect(Math.abs(width(narrow) - cssPxToEmu(430))).toBeLessThanOrEqual(1_000);
    expect(Math.abs(width(normal) - cssPxToEmu(600))).toBeLessThanOrEqual(1_000);
    expect(narrow).toContain('wrap="square"');
    expect(normal).toContain('wrap="square"');
    expect(explicit).toContain('<a:t>Erste Zeile</a:t>');
    expect(explicit).toContain('<a:t>Zweite Zeile</a:t>');
    expect((explicit.match(/<a:p>/g) || []).length).toBeGreaterThanOrEqual(2);
  }, 40_000);
});
