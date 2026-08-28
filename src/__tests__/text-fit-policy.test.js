import { beforeAll, describe, expect, it, vi } from 'vitest';
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

describe('editable text fit policy', () => {
  it('keeps the authored box geometry and asks PowerPoint to shrink only on overflow', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const label = document.createElement('div');
    label.textContent = 'Wirklogik';
    label.setAttribute(
      'style',
      'position:absolute;left:80px;top:800px;width:170px;height:48px;padding:8px 12px;' +
        'display:flex;align-items:center;white-space:nowrap;background:#0d1b2a;color:#fff;font-size:24px'
    );
    slide.appendChild(label);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    label.getBoundingClientRect = () => rect({ left: 80, top: 800, width: 170, height: 48 });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const shape = Array.from(doc.getElementsByTagName('p:sp')).find((candidate) =>
        Array.from(candidate.getElementsByTagName('a:t')).some((run) => run.textContent === 'Wirklogik')
      );

      expect(shape).toBeDefined();
      const bodyPr = shape.getElementsByTagName('a:bodyPr')[0];
      expect(bodyPr.getElementsByTagName('a:normAutofit')).toHaveLength(1);
      expect(bodyPr.getElementsByTagName('a:spAutoFit')).toHaveLength(0);
    } finally {
      slide.remove();
    }
  });

  it('maps native grid item centering to editable PowerPoint text alignment', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const badge = document.createElement('div');
    badge.textContent = '1';
    badge.setAttribute(
      'style',
      'position:absolute;left:80px;top:800px;width:68px;height:68px;' +
        'display:grid;align-items:center;justify-items:center;background:#2b6ff2;color:#fff;font-size:34px'
    );
    slide.appendChild(badge);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    badge.getBoundingClientRect = () => rect({ left: 80, top: 800, width: 68, height: 68 });
    const nativeGetComputedStyle = window.getComputedStyle.bind(window);
    const computedStyle = vi.spyOn(window, 'getComputedStyle').mockImplementation((element, pseudo) => {
      const style = nativeGetComputedStyle(element, pseudo);
      if (element !== badge || pseudo) return style;
      return new Proxy(style, {
        get(target, property) {
          if (property === 'display') return 'grid';
          if (property === 'alignItems' || property === 'justifyItems') return 'center';
          return Reflect.get(target, property);
        },
      });
    });

    try {
      const badgeStyle = window.getComputedStyle(badge);
      expect([badgeStyle.display, badgeStyle.alignItems, badgeStyle.justifyItems]).toEqual([
        'grid',
        'center',
        'center',
      ]);
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const shape = Array.from(doc.getElementsByTagName('p:sp')).find((candidate) =>
        Array.from(candidate.getElementsByTagName('a:t')).some((run) => run.textContent === '1')
      );

      expect(shape).toBeDefined();
      expect(shape.getElementsByTagName('a:bodyPr')[0].getAttribute('anchor')).toBe('ctr');
      expect(shape.getElementsByTagName('a:pPr')[0].getAttribute('algn')).toBe('ctr');
    } finally {
      computedStyle.mockRestore();
      slide.remove();
    }
  });

  it('renders ellipse decoration separately from its rectangular browser text box', async () => {
    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    const ellipse = document.createElement('div');
    ellipse.innerHTML = 'Einzelschritt<br>optimiert';
    ellipse.setAttribute(
      'style',
      'position:absolute;left:280px;top:500px;width:190px;height:102px;padding:20px 28px;' +
        'border:2px solid #096e66;border-radius:50%;font:700 22px/29px Arial;text-align:center;color:#096e66'
    );
    slide.appendChild(ellipse);
    document.body.appendChild(slide);

    slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: 1920, height: 1080 });
    ellipse.getBoundingClientRect = () => rect({ left: 280, top: 500, width: 190, height: 102 });

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      const doc = new DOMParser().parseFromString(xml, 'text/xml');
      const textShape = Array.from(doc.getElementsByTagName('p:sp')).find((candidate) =>
        Array.from(candidate.getElementsByTagName('a:t')).some((run) => run.textContent === 'Einzelschritt')
      );

      expect(textShape).toBeDefined();
      expect(textShape.getElementsByTagName('a:prstGeom')[0].getAttribute('prst')).toBe('rect');
      const ellipseShape = Array.from(doc.getElementsByTagName('p:sp')).find(
        (candidate) =>
          candidate !== textShape && candidate.getElementsByTagName('a:prstGeom')[0]?.getAttribute('prst') === 'ellipse'
      );
      expect(ellipseShape).toBeDefined();
      const bodyPr = textShape.getElementsByTagName('a:bodyPr')[0];
      expect(bodyPr.getElementsByTagName('a:normAutofit')).toHaveLength(1);
      expect(bodyPr.getElementsByTagName('a:spAutoFit')).toHaveLength(0);
      expect(Number(bodyPr.getAttribute('lIns'))).toBeGreaterThan(0);
      expect(Number(bodyPr.getAttribute('rIns'))).toBeGreaterThan(0);
      expect(Number(bodyPr.getAttribute('tIns'))).toBeGreaterThan(0);
      expect(Number(bodyPr.getAttribute('bIns'))).toBeGreaterThan(0);
      const titleRun = Array.from(textShape.getElementsByTagName('a:r')).find(
        (run) => run.getElementsByTagName('a:t')[0]?.textContent === 'Einzelschritt'
      );
      expect(titleRun.getElementsByTagName('a:rPr')[0].getAttribute('sz')).toBe('825');
    } finally {
      slide.remove();
    }
  });
});
