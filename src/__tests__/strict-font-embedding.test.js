import { beforeAll, describe, expect, it, vi } from 'vitest';
import { exportToPptx } from '../index.js';

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

describe('strict font embedding', () => {
  it('fails the export with structured details when a required font cannot be embedded', async () => {
    const slide = document.createElement('div');
    slide.textContent = 'Required font';
    document.body.appendChild(slide);
    slide.getBoundingClientRect = () => ({
      x: 0,
      y: 0,
      left: 0,
      top: 0,
      width: 1920,
      height: 1080,
      right: 1920,
      bottom: 1080,
      toJSON() {
        return this;
      },
    });

    const originalFetch = globalThis.fetch;
    globalThis.fetch = vi.fn(async () => ({ ok: false, status: 404 }));

    try {
      await expect(
        exportToPptx(slide, {
          skipDownload: true,
          skipNormalize: true,
          autoEmbedFonts: false,
          strictFontEmbedding: true,
          fonts: [{ name: 'Required Sans', weight: 400, url: 'https://example.invalid/required.ttf' }],
        })
      ).rejects.toMatchObject({
        code: 'DOM_TO_PPTX_FONT_EMBEDDING_FAILED',
        details: [expect.objectContaining({ label: 'Required Sans (regular)', ok: false })],
      });
    } finally {
      globalThis.fetch = originalFetch;
      slide.remove();
    }
  });
});
