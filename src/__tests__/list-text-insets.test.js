import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';
import { exportHtmlToPptx } from '../node-exporter.js';

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

// The marker width is measured in a hidden probe, which jsdom lays out as
// nothing. Answer for the probe only, so the arithmetic around the measurement
// can be tested without a browser.
async function withMeasuredMarkerHang(hangPx, run) {
  const original = Element.prototype.getBoundingClientRect;
  Element.prototype.getBoundingClientRect = function () {
    if (this.closest?.('[data-pptx-marker-probe]')) {
      const list = this.tagName === 'UL' || this.tagName === 'OL' ? this : this.closest('ul, ol');
      const marksInside = list?.style?.listStylePosition === 'inside';
      return rect({ left: this === list ? 0 : marksInside ? hangPx : 0, top: 0, width: 8, height: 8 });
    }
    return original.call(this);
  };
  try {
    return await run();
  } finally {
    Element.prototype.getBoundingClientRect = original;
  }
}

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

  it('collapses adjacent list-item margins instead of adding before and after spacing twice', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:1200px;top:220px;width:555px;height:174px;' +
        'font-size:19px;line-height:28px;margin:0;padding:0'
    );
    for (const text of ['First open point', 'Second open point', 'Third open point']) {
      const item = document.createElement('li');
      item.setAttribute('style', 'font-size:19px;line-height:28px;margin:17px 0;padding:0');
      item.textContent = text;
      list.appendChild(item);
    }
    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 1200, top: 220, width: 555, height: 174 });
    Array.from(list.children).forEach((item, index) => {
      item.getBoundingClientRect = () => rect({ left: 1200, top: 237 + index * 45, width: 555, height: 28 });
    });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const listShape = Array.from(doc.getElementsByTagName('p:sp')).find((shape) =>
        Array.from(shape.getElementsByTagName('a:t')).some((run) => run.textContent === 'First open point')
      );
      expect(listShape).toBeDefined();
      const paragraphs = Array.from(listShape.getElementsByTagName('a:p'));
      expect(paragraphs).toHaveLength(3);
      const spacing = paragraphs.map((paragraph) => ({
        before: paragraph.getElementsByTagName('a:spcBef').length,
        after: paragraph.getElementsByTagName('a:spcAft').length,
      }));

      expect(spacing).toEqual([
        { before: 1, after: 0 },
        { before: 1, after: 0 },
        { before: 1, after: 1 },
      ]);
    } finally {
      slide.remove();
    }
  });

  it('exports the empty-list boundary fixture as a valid slide', async () => {
    const documentNode = await emptyListDocument();
    expect(documentNode.getElementsByTagName('p:sld')).toHaveLength(1);
    expect(documentNode.getElementsByTagName('p:spTree')).toHaveLength(1);
  });

  it('preserves the browser-visible marker of an empty list item', async () => {
    const documentNode = await emptyListDocument();
    const bulletParagraphs = Array.from(documentNode.getElementsByTagName('a:p')).filter(
      (paragraph) => paragraph.getElementsByTagName('a:buChar').length === 1
    );
    expect(bulletParagraphs).toHaveLength(1);
  });

  it('assigns hierarchical indentLevel to indented sub-bullets instead of bloating hanging indent', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');

    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:500px;height:150px;color:#111;font-size:16px;line-height:25px;margin:0;padding:10px'
    );

    const liRoot = document.createElement('li');
    liRoot.textContent = 'Root bullet';
    const liSub = document.createElement('li');
    liSub.className = 'sub';
    liSub.textContent = 'Sub bullet indented';
    // Indent sub-bullet by 24px (1 indent level)
    liSub.setAttribute('style', 'margin-left: 24px;');

    list.appendChild(liRoot);
    list.appendChild(liSub);
    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 500, height: 150 });
    // Root li at x=210 (due to ul padding 10)
    liRoot.getBoundingClientRect = () => rect({ left: 210, top: 210, width: 480, height: 25 });
    // Sub li shifted right by 24px -> x=234
    liSub.getBoundingClientRect = () => rect({ left: 234, top: 235, width: 456, height: 25 });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');

      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const paragraphs = Array.from(doc.getElementsByTagName('a:p'));

      const subP = paragraphs.find((p) =>
        Array.from(p.getElementsByTagName('a:t')).some((t) => t.textContent.includes('Sub bullet indented'))
      );
      expect(subP).toBeDefined();

      const pPr = subP.getElementsByTagName('a:pPr')[0];
      expect(pPr).toBeDefined();

      // lvl attribute must be 1 for sub-bullet
      expect(pPr.getAttribute('lvl')).toBe('1');

      // indent should be negative standard hanging indent (-10pt scaled = -127000 EMU)
      // and marL should be indentLevel adjusted (marL + marL * lvl)
      const marL = Number(pPr.getAttribute('marL'));
      const indent = Number(pPr.getAttribute('indent'));
      expect(indent).toBeLessThan(0);
      expect(marL).toBeGreaterThan(0);
    } finally {
      slide.remove();
    }
  });

  // The browser paints an outside marker in the list's left padding and starts
  // the text at the content edge. PowerPoint ties glyph and text to one number,
  // so the room for the marker has to come out of the inset -- added on top of
  // it, every item's text sat one marker width too far right.
  it('takes the measured marker width out of the list inset instead of adding it to the text', async () => {
    const paddingLeftPx = 24;
    const hangPx = 20;

    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:400px;height:60px;color:#111;font-size:21px;' +
        `line-height:30px;margin:0;padding:0 0 0 ${paddingLeftPx}px`
    );
    const item = document.createElement('li');
    item.setAttribute('style', 'font-size:21px;line-height:30px');
    item.textContent = 'Measured item';
    list.appendChild(item);
    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 400, height: 60 });
    item.getBoundingClientRect = () => rect({ left: 200 + paddingLeftPx, top: 200, width: 376, height: 30 });

    try {
      const doc = await withMeasuredMarkerHang(hangPx, async () => {
        const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
        const zip = await JSZip.loadAsync(blob);
        return new DOMParser().parseFromString(await zip.file('ppt/slides/slide1.xml').async('string'), 'text/xml');
      });
      const listShape = Array.from(doc.getElementsByTagName('p:sp')).find((shape) =>
        Array.from(shape.getElementsByTagName('a:t')).some((run) => run.textContent === 'Measured item')
      );
      expect(listShape).toBeDefined();

      const emu = (px) => Math.round(px * PT_PER_PX * EMU_PER_PT);
      const lIns = Number(listShape.getElementsByTagName('a:bodyPr')[0].getAttribute('lIns'));
      const pPr = listShape.getElementsByTagName('a:pPr')[0];
      const marL = Number(pPr.getAttribute('marL'));
      const indent = Number(pPr.getAttribute('indent'));

      // The text sits at the content edge, the glyph one marker width left of it.
      expect(lIns + marL).toBe(emu(paddingLeftPx));
      expect(lIns + marL + indent).toBe(emu(paddingLeftPx - hangPx));
      expect(pPr.getAttribute('lvl')).toBeNull();
    } finally {
      slide.remove();
    }
  });

  // An item without a marker cannot carry a hanging indent: PptxGenJS pins such
  // a paragraph to marL=0. A list that mixes the two would tear apart, so the
  // inset stays where it was.
  it('leaves the list inset alone when an item has no marker', async () => {
    const paddingLeftPx = 24;

    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const list = document.createElement('ul');
    list.setAttribute(
      'style',
      'position:absolute;left:200px;top:200px;width:400px;height:60px;color:#111;font-size:22px;' +
        `line-height:30px;margin:0;padding:0 0 0 ${paddingLeftPx}px`
    );
    const marked = document.createElement('li');
    marked.setAttribute('style', 'font-size:22px;line-height:30px');
    marked.textContent = 'Marked item';
    const bare = document.createElement('li');
    bare.setAttribute('style', 'font-size:22px;line-height:30px;list-style-type:none');
    bare.textContent = 'Bare item';
    list.append(marked, bare);
    slide.appendChild(list);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    list.getBoundingClientRect = () => rect({ left: 200, top: 200, width: 400, height: 60 });
    Array.from(list.children).forEach((child, index) => {
      child.getBoundingClientRect = () =>
        rect({ left: 200 + paddingLeftPx, top: 200 + index * 30, width: 376, height: 30 });
    });

    try {
      const doc = await withMeasuredMarkerHang(20, async () => {
        const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
        const zip = await JSZip.loadAsync(blob);
        return new DOMParser().parseFromString(await zip.file('ppt/slides/slide1.xml').async('string'), 'text/xml');
      });
      const listShape = Array.from(doc.getElementsByTagName('p:sp')).find((shape) =>
        Array.from(shape.getElementsByTagName('a:t')).some((run) => run.textContent === 'Marked item')
      );
      expect(listShape).toBeDefined();
      expect(Number(listShape.getElementsByTagName('a:bodyPr')[0].getAttribute('lIns'))).toBe(
        Math.round(paddingLeftPx * PT_PER_PX * EMU_PER_PT)
      );
    } finally {
      slide.remove();
    }
  });
});

// A list item that paints -- a separator line, a dot drawn by its ::before --
// has nowhere to go in one text box for the whole list. The list takes the
// ordinary path then, which draws both; pseudo-elements are no DOM children,
// so the out-of-flow check alone never saw the dot.
describe('lists whose items paint', () => {
  const exportList = async (itemCss, beforeCss) => {
    const html = `<!doctype html><html><head><style>
        * { box-sizing: border-box; margin: 0; }
        .slide { position: relative; width: 1280px; height: 720px; background: white; }
        ul { position: absolute; left: 100px; top: 100px; width: 360px; list-style: none; padding: 0; }
        li { font: 14pt/20pt Arial, sans-serif; color: #4B5563; ${itemCss} }
        li::before { ${beforeCss} }
      </style></head><body><section class="slide"><ul><li>Erster Punkt</li><li>Zweiter Punkt</li><li>Dritter Punkt</li></ul></section></body></html>`;
    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false },
    });
    const xml = await (await JSZip.loadAsync(buffer)).file('ppt/slides/slide1.xml').async('string');
    return Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (match) => match[0]);
  };

  it('keeps the dot markers and separator lines of painted items', async () => {
    const shapes = await exportList(
      'position: relative; padding: 10px 0 10px 20px; border-bottom: 1px solid #E5E7EB;',
      'content: ""; position: absolute; left: 0; top: 18px; width: 7px; height: 7px; background: #F24726; border-radius: 50%;'
    );
    const withText = (text) => shapes.filter((shape) => shape.includes(`<a:t>${text}</a:t>`));

    for (const item of ['Erster Punkt', 'Zweiter Punkt', 'Dritter Punkt']) expect(withText(item)).toHaveLength(1);
    expect(shapes.filter((shape) => shape.includes('prst="ellipse"') && shape.includes('val="F24726"'))).toHaveLength(
      3
    );
    expect(shapes.filter((shape) => shape.includes('val="E5E7EB"') && !shape.includes('<a:t>'))).toHaveLength(3);
  });

  it('keeps a list whose markers are inline text in one text box', async () => {
    const shapes = await exportList('', 'content: "– ";');
    const listShapes = shapes.filter((shape) => shape.includes('Punkt</a:t>'));

    expect(listShapes).toHaveLength(1);
    expect(listShapes[0]).toContain('<a:t>Dritter Punkt</a:t>');
  });
});
