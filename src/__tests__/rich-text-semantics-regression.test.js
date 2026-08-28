import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';

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

function runProperties(documentNode, text) {
  const textNode = Array.from(documentNode.getElementsByTagName('a:t')).find((node) => node.textContent === text);
  expect(textNode).toBeDefined();
  const run = textNode.parentElement;
  expect(run.tagName).toBe('a:r');
  return run.getElementsByTagName('a:rPr')[0];
}

beforeAll(() => {
  let fillStyle = '';
  HTMLCanvasElement.prototype.getContext = () => ({
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value) {
      fillStyle = value;
    },
    clearRect: () => {},
    fillRect: () => {},
    getImageData: () => ({ data: [0, 0, 0, 255] }),
  });
});

afterEach(() => {
  document.body.replaceChildren();
});

describe('editable rich-text semantics', () => {
  it('maps line-through, superscript and subscript to DrawingML run properties', async () => {
    const slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');

    const text = document.createElement('div');
    text.setAttribute(
      'style',
      'position:absolute;left:160px;top:180px;width:1200px;height:120px;' + 'font:40px/1.3 Arial;color:#101828'
    );
    text.innerHTML =
      '<span style="text-decoration-line:line-through">DEPRECATED</span> ' +
      '<span style="vertical-align:super">SUPER</span> ' +
      '<span style="vertical-align:sub">SUB</span>';

    slide.appendChild(text);
    document.body.appendChild(slide);
    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    text.getBoundingClientRect = () => rect({ left: 160, top: 180, width: 1200, height: 120 });

    const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
    const zip = await JSZip.loadAsync(blob);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const documentNode = new DOMParser().parseFromString(xml, 'text/xml');

    expect(runProperties(documentNode, 'DEPRECATED').getAttribute('strike')).toBe('sngStrike');
    expect(runProperties(documentNode, 'SUPER').getAttribute('baseline')).toBe('30000');
    expect(runProperties(documentNode, 'SUB').getAttribute('baseline')).toBe('-40000');
  });
});
