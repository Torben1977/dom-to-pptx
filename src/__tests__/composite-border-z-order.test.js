import { beforeAll, describe, expect, it } from 'vitest';
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

describe('composite border z-order', () => {
  it('renders an asymmetric border above a filled text shape', async () => {
    const slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background-color:#fff');

    const card = document.createElement('div');
    card.textContent = 'Working hypothesis';
    card.setAttribute(
      'style',
      'position:absolute;left:120px;top:220px;width:900px;height:180px;' +
        'box-sizing:border-box;padding:24px;background-color:#FFF4D6;color:#102A43;' +
        'border-top:1px solid #D9E2EA;border-right:1px solid #D9E2EA;' +
        'border-bottom:1px solid #D9E2EA;border-left:8px solid #E2A72E;font-size:32px'
    );
    slide.appendChild(card);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    card.getBoundingClientRect = () => rect({ left: 120, top: 220, width: 900, height: 180 });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const documentNode = new DOMParser().parseFromString(xml, 'text/xml');
      const shapeTree = documentNode.getElementsByTagName('p:spTree')[0];
      const drawingNodes = Array.from(shapeTree.children).filter((node) => ['p:sp', 'p:pic'].includes(node.tagName));
      const cardIndex = drawingNodes.findIndex((node) =>
        Array.from(node.getElementsByTagName('a:t')).some((text) => text.textContent.includes('Working hypothesis'))
      );
      const borderIndex = drawingNodes.findIndex((node) => node.tagName === 'p:pic');

      expect(cardIndex).toBeGreaterThanOrEqual(0);
      expect(borderIndex).toBeGreaterThan(cardIndex);
    } finally {
      slide.remove();
    }
  });

  it('renders asymmetric borders above a gradient background', async () => {
    const slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background-color:#fff');

    const matrix = document.createElement('div');
    matrix.setAttribute(
      'style',
      'position:absolute;left:280px;top:270px;width:1130px;height:570px;' +
        'background:linear-gradient(135deg,rgba(6,93,201,.05),rgba(9,110,102,.10));' +
        'border-left:3px solid #0E1B2C;border-bottom:3px solid #0E1B2C'
    );
    slide.appendChild(matrix);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    matrix.getBoundingClientRect = () => rect({ left: 280, top: 270, width: 1130, height: 570 });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const documentNode = new DOMParser().parseFromString(xml, 'text/xml');
      const shapeTree = documentNode.getElementsByTagName('p:spTree')[0];
      const drawingNodes = Array.from(shapeTree.children).filter((node) => ['p:sp', 'p:pic'].includes(node.tagName));
      const gradientIndex = drawingNodes.findIndex((node) => node.tagName === 'p:pic');
      const borderIndices = drawingNodes
        .map((node, index) => ({ node, index }))
        .filter(({ node }) => {
          if (node.tagName !== 'p:sp' || node.getElementsByTagName('a:t').length > 0) return false;
          return Array.from(node.getElementsByTagName('a:srgbClr')).some(
            (color) => color.getAttribute('val') === '0E1B2C'
          );
        })
        .map(({ index }) => index);

      expect(gradientIndex).toBeGreaterThanOrEqual(0);
      expect(borderIndices).toHaveLength(2);
      expect(borderIndices.every((index) => index > gradientIndex)).toBe(true);
    } finally {
      slide.remove();
    }
  });
});
