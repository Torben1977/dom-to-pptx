import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';

// A 1920px-wide root is exported as a 10in slide, so 1 CSS px is 914400 / 96 / 2 EMU,
// and 1 CSS px of padding is 0.75 / 2 pt.
const EMU_PER_PT = 12700;
const PT_PER_PX = 0.75 * 0.5;

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

let emptyListDocumentPromise;

function emptyListDocument() {
  if (emptyListDocumentPromise) return emptyListDocumentPromise;

  emptyListDocumentPromise = (async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:500px;height:80px;' +
        'font-size:24px;line-height:32px;margin:0;padding:0 0 0 24px'
    );
    const item = document.createElement('li');
    list.appendChild(item);
    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 500, height: 80 });
    item.getBoundingClientRect = () => rect({ left: 224, top: 200, width: 476, height: 32 });

    try {
      const blob = await exportToPptx(slide, {
        skipDownload: true,
        autoEmbedFonts: false,
        skipNormalize: true,
      });
      const zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      return new DOMParser().parseFromString(xml, 'text/xml');
    } finally {
      slide.remove();
    }
  })();

  return emptyListDocumentPromise;
}

describe('list text insets', () => {
  it('maps each CSS padding edge of a <ul> to the matching PPTX inset', async () => {
    // Four distinct paddings, so a swapped pair cannot go unnoticed.
    const paddingPx = { top: 12, right: 4, bottom: 8, left: 20 };

    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');

    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:400px;height:120px;color:#111;font-size:16px;' +
        `line-height:25px;margin:0;padding:${paddingPx.top}px ${paddingPx.right}px ${paddingPx.bottom}px ${paddingPx.left}px`
    );
    for (const text of ['First item', 'Second item']) {
      const item = document.createElement('li');
      item.setAttribute('style', 'font-size:16px;line-height:25px');
      item.textContent = text;
      list.appendChild(item);
    }

    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 400, height: 120 });
    Array.from(list.children).forEach((item, index) => {
      item.getBoundingClientRect = () => rect({ left: 220, top: 212 + index * 25, width: 376, height: 25 });
    });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');

      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const listShape = Array.from(doc.getElementsByTagName('p:sp')).find((shape) =>
        Array.from(shape.getElementsByTagName('a:t')).some((run) => run.textContent.includes('First item'))
      );
      expect(listShape).toBeDefined();

      const bodyPr = listShape.getElementsByTagName('a:bodyPr')[0];
      const insetEmu = (px) => Math.round(px * PT_PER_PX * EMU_PER_PT);

      expect(Number(bodyPr.getAttribute('lIns'))).toBe(insetEmu(paddingPx.left));
      expect(Number(bodyPr.getAttribute('rIns'))).toBe(insetEmu(paddingPx.right));
      expect(Number(bodyPr.getAttribute('bIns'))).toBe(insetEmu(paddingPx.bottom));
      expect(Number(bodyPr.getAttribute('tIns'))).toBe(insetEmu(paddingPx.top));
    } finally {
      slide.remove();
    }
  });

  it('preserves ordered-list start and explicit item values', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const list = document.createElement('ol');
    list.start = 2;
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:400px;height:120px;color:#111;font-size:16px;' +
        'line-height:25px;margin:0;padding:0 0 0 24px'
    );
    for (const [index, text] of ['First item', 'Second item', 'Third item'].entries()) {
      const item = document.createElement('li');
      item.setAttribute('style', 'font-size:16px;line-height:25px');
      if (index === 1) item.value = 5;
      item.textContent = text;
      list.appendChild(item);
    }
    slide.appendChild(list);
    document.body.appendChild(slide);
    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 400, height: 120 });
    Array.from(list.children).forEach((item, index) => {
      item.getBoundingClientRect = () => rect({ left: 224, top: 200 + index * 25, width: 376, height: 25 });
    });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const autoNumbers = Array.from(doc.getElementsByTagName('a:buAutoNum'));

      expect(autoNumbers).toHaveLength(3);
      expect(autoNumbers.map((item) => item.getAttribute('startAt'))).toEqual(['2', '5', '6']);
    } finally {
      slide.remove();
    }
  });

  it('emits one bullet definition for a rich-text list item', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:500px;height:120px;' +
        'font-size:24px;line-height:32px;margin:0;padding:0 0 0 24px'
    );
    list.innerHTML = '<li><strong>Bold lead</strong> normal continuation</li>';
    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 500, height: 120 });
    list.firstElementChild.getBoundingClientRect = () => rect({ left: 224, top: 200, width: 476, height: 32 });

    try {
      const blob = await exportToPptx(slide, {
        skipDownload: true,
        autoEmbedFonts: false,
        skipNormalize: true,
      });
      const zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const listShape = Array.from(doc.getElementsByTagName('p:sp')).find((shape) =>
        Array.from(shape.getElementsByTagName('a:t')).some((run) => run.textContent === 'Bold lead')
      );

      expect(listShape).toBeDefined();
      expect(Array.from(listShape.getElementsByTagName('a:t')).map((run) => run.textContent)).toEqual([
        'Bold lead',
        ' normal continuation',
      ]);
      expect(listShape.getElementsByTagName('a:p')).toHaveLength(1);
      expect(listShape.getElementsByTagName('a:buChar')).toHaveLength(1);
    } finally {
      slide.remove();
    }
  });

  it('exports the empty-list boundary fixture as a valid slide', async () => {
    const documentNode = await emptyListDocument();
    expect(documentNode.getElementsByTagName('p:sld')).toHaveLength(1);
    expect(documentNode.getElementsByTagName('p:spTree')).toHaveLength(1);
  });

  it.fails('preserves the browser-visible marker of an empty list item', async () => {
    const documentNode = await emptyListDocument();
    const bulletParagraphs = Array.from(documentNode.getElementsByTagName('a:p')).filter(
      (paragraph) => paragraph.getElementsByTagName('a:buChar').length === 1
    );
    expect(bulletParagraphs).toHaveLength(1);
  });
});
