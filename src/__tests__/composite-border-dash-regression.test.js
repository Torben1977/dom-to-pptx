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

function nativeBorderHasDash(documentNode, color, acceptedValues) {
  return Array.from(documentNode.getElementsByTagName('p:sp')).some((shape) => {
    const hasColor = Array.from(shape.getElementsByTagName('a:srgbClr')).some(
      (entry) => entry.getAttribute('val') === color
    );
    const dash = shape.getElementsByTagName('a:prstDash')[0]?.getAttribute('val');
    return hasColor && acceptedValues.includes(dash);
  });
}

function svgBorderHasDash(media, color, { dotted = false } = {}) {
  return media.some((svg) => {
    if (!svg.toUpperCase().includes(`#${color}`)) return false;
    if (!svg.includes('stroke-dasharray')) return false;
    return !dotted || /stroke-linecap=["']round["']/.test(svg);
  });
}

function transformOf(node) {
  const transform = node.getElementsByTagName('a:xfrm')[0];
  const offset = transform.getElementsByTagName('a:off')[0];
  const extent = transform.getElementsByTagName('a:ext')[0];
  return {
    x: Number(offset.getAttribute('x')),
    y: Number(offset.getAttribute('y')),
    w: Number(extent.getAttribute('cx')),
    h: Number(extent.getAttribute('cy')),
  };
}

function shapeWithStroke(documentNode, color) {
  return Array.from(documentNode.getElementsByTagName('p:sp')).find((shape) => {
    const line = shape.getElementsByTagName('a:ln')[0];
    return (
      line && Array.from(line.getElementsByTagName('a:srgbClr')).some((entry) => entry.getAttribute('val') === color)
    );
  });
}

describe('composite border dash semantics', () => {
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

  async function exportBorder(borderStyle, color) {
    const slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const card = document.createElement('div');
    card.setAttribute(
      'style',
      `position:absolute;left:120px;top:220px;width:900px;height:180px;` +
        `box-sizing:border-box;background:#eef2f6;border-left:8px ${borderStyle} #${color}`
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
      const media = [];
      for (const name of Object.keys(zip.files)) {
        if (!name.startsWith('ppt/media/') || !name.endsWith('.svg')) continue;
        media.push(await zip.file(name).async('string'));
      }
      return { documentNode, media };
    } finally {
      slide.remove();
    }
  }

  it('keeps a one-sided dashed border dashed in the PowerPoint output', async () => {
    const { documentNode, media } = await exportBorder('dashed', 'F04444');
    const retainedAsNativeLine = nativeBorderHasDash(documentNode, 'F04444', ['dash', 'lgDash', 'sysDash']);
    const retainedAsSvg = svgBorderHasDash(media, 'F04444');
    expect(retainedAsNativeLine || retainedAsSvg).toBe(true);
  });

  it('keeps a one-sided dotted border dotted in the PowerPoint output', async () => {
    const { documentNode, media } = await exportBorder('dotted', '008A7A');
    const retainedAsNativeLine = nativeBorderHasDash(documentNode, '008A7A', ['dot', 'sysDot', 'roundDot']);
    const retainedAsSvg = svgBorderHasDash(media, '008A7A', { dotted: true });
    expect(retainedAsNativeLine || retainedAsSvg).toBe(true);
  });

  it('centers native dashed and dotted borders inside a gradient border box', async () => {
    const slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const card = document.createElement('div');
    card.setAttribute(
      'style',
      'position:absolute;left:120px;top:220px;width:900px;height:180px;' +
        'box-sizing:border-box;background:linear-gradient(90deg,#fff,#eef2f6);' +
        'border-left:8px dashed #F04444;border-top:6px dotted #12A150;' +
        'border-right:10px dashed #3157D5;border-bottom:12px dotted #B65C00'
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
      const background = documentNode.getElementsByTagName('p:pic')[0];
      const box = transformOf(background);

      const left = shapeWithStroke(documentNode, 'F04444');
      const top = shapeWithStroke(documentNode, '12A150');
      const right = shapeWithStroke(documentNode, '3157D5');
      const bottom = shapeWithStroke(documentNode, 'B65C00');
      expect(left && top && right && bottom).toBeTruthy();

      const leftWidth = Number(left.getElementsByTagName('a:ln')[0].getAttribute('w'));
      const topWidth = Number(top.getElementsByTagName('a:ln')[0].getAttribute('w'));
      const rightWidth = Number(right.getElementsByTagName('a:ln')[0].getAttribute('w'));
      const bottomWidth = Number(bottom.getElementsByTagName('a:ln')[0].getAttribute('w'));

      expect(transformOf(left).x).toBe(Math.round(box.x + leftWidth / 2));
      expect(transformOf(top).y).toBe(Math.round(box.y + topWidth / 2));
      expect(transformOf(right).x).toBe(Math.round(box.x + box.w - rightWidth / 2));
      expect(transformOf(bottom).y).toBe(Math.round(box.y + box.h - bottomWidth / 2));
    } finally {
      slide.remove();
    }
  });
});
