import { afterAll, beforeAll, describe, expect, it } from 'vitest';
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

function findTextShape(documentNode, text) {
  return Array.from(documentNode.getElementsByTagName('p:sp')).find((shape) =>
    Array.from(shape.getElementsByTagName('a:t')).some((run) => run.textContent === text)
  );
}

describe('logical text alignment', () => {
  let slide;
  let documentNode;

  beforeAll(async () => {
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

    slide = document.createElement('section');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');

    const cases = [
      {
        text: 'ROW_REVERSE_START',
        style:
          'left:100px;top:100px;display:flex;flex-direction:row-reverse;' +
          'justify-content:flex-start;align-items:center',
      },
      {
        text: 'ROW_REVERSE_LOGICAL_START',
        style:
          'left:100px;top:220px;display:flex;flex-direction:row-reverse;' + 'justify-content:start;align-items:center',
      },
      {
        text: 'RTL_ROW_REVERSE_START',
        style:
          'left:100px;top:340px;display:flex;flex-direction:row-reverse;direction:rtl;' +
          'justify-content:flex-start;align-items:center',
      },
      {
        text: 'RTL_ROW_REVERSE_LOGICAL_START',
        style:
          'left:100px;top:460px;display:flex;flex-direction:row-reverse;direction:rtl;' +
          'justify-content:start;align-items:center',
      },
      {
        text: 'COLUMN_REVERSE_START',
        style:
          'left:800px;top:100px;display:flex;flex-direction:column-reverse;' +
          'justify-content:flex-start;align-items:flex-start',
      },
      {
        text: 'COLUMN_REVERSE_LOGICAL_START',
        style:
          'left:800px;top:260px;display:flex;flex-direction:column-reverse;' +
          'justify-content:start;align-items:flex-start',
      },
      {
        text: 'RTL_START',
        style: 'left:800px;top:420px;direction:rtl;text-align:start',
      },
      {
        text: 'RTL_END',
        style: 'left:800px;top:580px;direction:rtl;text-align:end',
      },
      {
        text: 'LTR_START_BASELINE',
        style: 'left:800px;top:740px;direction:ltr;text-align:start',
      },
    ];

    for (const entry of cases) {
      const box = document.createElement('div');
      box.textContent = entry.text;
      box.setAttribute(
        'style',
        `position:absolute;width:600px;height:120px;${entry.style};` +
          'box-sizing:border-box;background:#eef2f6;font:32px/1.2 Arial'
      );
      box.getBoundingClientRect = () => {
        const left = Number.parseFloat(box.style.left);
        const top = Number.parseFloat(box.style.top);
        return rect({ left, top, width: 600, height: 120 });
      };
      slide.appendChild(box);
    }

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    document.body.appendChild(slide);

    const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
    const zip = await JSZip.loadAsync(blob);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    documentNode = new DOMParser().parseFromString(xml, 'text/xml');
  });

  afterAll(() => {
    slide?.remove();
  });

  it('maps row-reverse flex-start to the visual right edge', () => {
    const shape = findTextShape(documentNode, 'ROW_REVERSE_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('r');
  });

  it('keeps logical start at the LTR inline start for row-reverse', () => {
    const shape = findTextShape(documentNode, 'ROW_REVERSE_LOGICAL_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('l');
  });

  it('maps RTL row-reverse flex-start to the reversed flex edge', () => {
    const shape = findTextShape(documentNode, 'RTL_ROW_REVERSE_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('l');
  });

  it('keeps logical start at the RTL inline start for row-reverse', () => {
    const shape = findTextShape(documentNode, 'RTL_ROW_REVERSE_LOGICAL_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('r');
  });

  it('maps column-reverse flex-start to the visual bottom edge', () => {
    const shape = findTextShape(documentNode, 'COLUMN_REVERSE_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:bodyPr')[0].getAttribute('anchor')).toBe('b');
  });

  it('keeps logical start at the block start for column-reverse', () => {
    const shape = findTextShape(documentNode, 'COLUMN_REVERSE_LOGICAL_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:bodyPr')[0].getAttribute('anchor')).toBe('t');
  });

  it('resolves text-align start against an RTL writing direction', () => {
    const shape = findTextShape(documentNode, 'RTL_START');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('r');
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('rtl')).toBe('1');
  });

  it('resolves text-align end against an RTL writing direction', () => {
    const shape = findTextShape(documentNode, 'RTL_END');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('l');
  });

  it('keeps the ordinary LTR start alignment unchanged', () => {
    const shape = findTextShape(documentNode, 'LTR_START_BASELINE');
    expect(shape).toBeDefined();
    expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('l');
  });
});
