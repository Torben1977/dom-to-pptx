import { describe, it, expect } from 'vitest';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { createShapeMargin, createTableCellMargin } from '../utils.js';

describe('margin helpers', () => {
  it('createShapeMargin normalizes CSS (top, right, bottom, left) to PptxGenJS [lIns, rIns, bIns, tIns]', () => {
    const top = 10;
    const right = 20;
    const bottom = 30;
    const left = 40;

    const shapeMargin = createShapeMargin(top, right, bottom, left);
    expect(shapeMargin).toEqual([40, 20, 30, 10]); // [left, right, bottom, top]
  });

  it('createTableCellMargin keeps CSS order and hands PptxGenJS inches', () => {
    expect(createTableCellMargin(9, 18, 36, 72)).toEqual([0.125, 0.25, 0.5, 1]);
  });

  // PptxGenJS reads a cell margin as points only when its first value is >= 1.
  // A sub-point top padding must not flip the unit — the default 1px cell padding
  // came out as 0.75 inch per side.
  it.each([
    ['browser default 1px padding', 0.75],
    ['no padding', 0],
    ['6px padding', 4.5],
  ])('keeps the cell margin in points for %s', async (_, pt) => {
    const pptx = new PptxGenJS();
    pptx.addSlide().addTable([[{ text: 'Zelle', options: { margin: createTableCellMargin(pt, pt, pt, pt) } }]], {
      x: 0,
      y: 0,
      w: 4,
    });
    const zip = await JSZip.loadAsync(await pptx.write({ outputType: 'nodebuffer' }));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const emu = Math.round(pt * 12700);
    expect(xml).toContain(`marL="${emu}" marR="${emu}" marT="${emu}" marB="${emu}"`);
  });
});
