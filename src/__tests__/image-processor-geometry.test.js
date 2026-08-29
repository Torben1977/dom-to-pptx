import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  getProcessedImage,
  normalizeImageCornerRadii,
  resolveObjectPositionOffset,
} from '../image-processor.js';

describe('image processing geometry', () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('resolves computed calc object-position without losing the pixel offset', () => {
    expect(resolveObjectPositionOffset(
      'calc(100% - 20px) calc(100% - 10px)',
      200,
      100,
      400,
      160
    )).toEqual({ x: -220, y: -70 });
    expect(resolveObjectPositionOffset('25% 75%', 200, 100, 400, 160)).toEqual({ x: -50, y: -45 });
    expect(resolveObjectPositionOffset('top', 200, 100, 400, 160)).toEqual({ x: -100, y: 0 });
    expect(resolveObjectPositionOffset('right 20px bottom 10px', 200, 100, 400, 160)).toEqual({
      x: -220,
      y: -70,
    });
  });

  it('keeps independent horizontal and vertical corner radii', () => {
    expect(normalizeImageCornerRadii({
      tl: { x: 100, y: 20 },
      tr: { x: 50, y: 40 },
      br: { x: 0, y: 0 },
      bl: { x: 10, y: 30 },
    }, 200, 100)).toEqual({
      tl: { x: 100, y: 20 },
      tr: { x: 50, y: 40 },
      br: { x: 0, y: 0 },
      bl: { x: 10, y: 30 },
    });
  });

  it('uses the resolved crop and elliptical radii in the canvas processing path', async () => {
    const context = {
      scale: vi.fn(),
      beginPath: vi.fn(),
      moveTo: vi.fn(),
      lineTo: vi.fn(),
      ellipse: vi.fn(),
      closePath: vi.fn(),
      fill: vi.fn(),
      drawImage: vi.fn(),
      fillStyle: '',
      globalCompositeOperation: '',
    };
    const canvas = {
      width: 0,
      height: 0,
      getContext: () => context,
      toDataURL: () => 'data:image/png;base64,processed',
    };
    const originalCreateElement = document.createElement.bind(document);
    vi.spyOn(document, 'createElement').mockImplementation((tagName, options) => (
      tagName === 'canvas' ? canvas : originalCreateElement(tagName, options)
    ));
    class FakeImage {
      width = 400;
      height = 160;
      crossOrigin = '';
      onload = null;
      onerror = null;
      set src(_value) {
        this.onload?.();
      }
    }
    vi.stubGlobal('Image', FakeImage);

    const result = await getProcessedImage(
      'data:image/png;base64,source',
      200,
      100,
      {
        tl: { x: 100, y: 20 },
        tr: { x: 50, y: 40 },
        br: { x: 0, y: 0 },
        bl: { x: 10, y: 30 },
      },
      'none',
      'calc(100% - 20px) calc(100% - 10px)'
    );

    expect(result).toBe('data:image/png;base64,processed');
    expect(context.drawImage).toHaveBeenCalledWith(expect.any(FakeImage), -220, -70, 400, 160);
    expect(context.ellipse).toHaveBeenCalledWith(100, 20, 100, 20, 0, Math.PI, (3 * Math.PI) / 2);
    expect(context.ellipse).toHaveBeenCalledWith(150, 40, 50, 40, 0, -Math.PI / 2, 0);
  });
});
