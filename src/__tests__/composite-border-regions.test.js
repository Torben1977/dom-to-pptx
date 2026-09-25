import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';

// A solid side of a composite border is the band the browser fills: cut
// diagonally where it meets a wide neighbour, a plain rectangle beside a
// zero-width one, and a triangle in a box without content — the CSS arrow,
// which used to arrive in PowerPoint as a filled rectangle.

function rect({ left, top, width, height }) {
  return {
    x: left,
    y: top,
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    toJSON() {
      return this;
    },
  };
}

function shapesFilled(documentNode, color) {
  return Array.from(documentNode.getElementsByTagName('p:sp')).filter((shape) =>
    Array.from(shape.getElementsByTagName('a:solidFill')).some(
      (fill) => fill.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val') === color
    )
  );
}

/** The outline of a custom shape as fractions of its path box, distinct corners only. */
function outline(shape) {
  const path = shape.getElementsByTagName('a:path')[0];
  const [w, h] = [Number(path.getAttribute('w')), Number(path.getAttribute('h'))];
  const corners = Array.from(path.getElementsByTagName('a:pt'), (point) => [
    Math.round((Number(point.getAttribute('x')) / w) * 100) / 100,
    Math.round((Number(point.getAttribute('y')) / h) * 100) / 100,
  ]);
  return corners.filter(
    (corner, index) => corners.findIndex((other) => other[0] === corner[0] && other[1] === corner[1]) === index
  );
}

describe('composite border regions', () => {
  beforeAll(() => {
    HTMLCanvasElement.prototype.getContext = () => ({
      fillStyle: '',
      clearRect: () => {},
      fillRect: () => {},
      getImageData: () => ({ data: [0, 0, 0, 255] }),
    });
  });

  async function exportBox(style, box) {
    const slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const element = document.createElement('div');
    element.setAttribute('style', `position:absolute;left:${box.left}px;top:${box.top}px;${style}`);
    slide.appendChild(element);
    document.body.appendChild(slide);
    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    element.getBoundingClientRect = () => rect(box);
    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      return new DOMParser().parseFromString(xml, 'text/xml');
    } finally {
      slide.remove();
    }
  }

  it('draws a CSS arrow as a native triangle and leaves out the transparent sides', async () => {
    const documentNode = await exportBox(
      'width:0;height:0;border-top:22px solid transparent;border-bottom:22px solid transparent;' +
        'border-left:30px solid #0F766E',
      { left: 400, top: 300, width: 30, height: 44 }
    );

    const shapes = shapesFilled(documentNode, '0F766E');
    expect(shapes).toHaveLength(1);
    expect(shapes[0].getElementsByTagName('a:custGeom')).toHaveLength(1);
    expect(outline(shapes[0])).toEqual([
      [0, 1],
      [0, 0],
      [1, 0.5],
    ]);
    expect(shapesFilled(documentNode, '000000')).toEqual([]);
  });

  it('keeps a one-sided accent a rectangle', async () => {
    const documentNode = await exportBox(
      'width:300px;height:120px;box-sizing:border-box;border-left:8px solid #B45309',
      {
        left: 100,
        top: 100,
        width: 300,
        height: 120,
      }
    );

    const shapes = shapesFilled(documentNode, 'B45309');
    expect(shapes).toHaveLength(1);
    expect(shapes[0].getElementsByTagName('a:prstGeom')[0]?.getAttribute('prst')).toBe('rect');
  });

  it('mitres two wide sides of different colours where they meet', async () => {
    const documentNode = await exportBox(
      'width:100px;height:100px;box-sizing:border-box;border-top:20px solid #C0392B;border-left:20px solid #2E86C1',
      { left: 100, top: 100, width: 100, height: 100 }
    );

    const [top] = shapesFilled(documentNode, 'C0392B');
    const [left] = shapesFilled(documentNode, '2E86C1');
    expect(outline(top)).toEqual([
      [0, 0],
      [1, 0],
      [1, 1],
      [0.2, 1],
    ]);
    expect(outline(left)).toEqual([
      [0, 1],
      [0, 0],
      [1, 0.2],
      [1, 1],
    ]);
  });
});
