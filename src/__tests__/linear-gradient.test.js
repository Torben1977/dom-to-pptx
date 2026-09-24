import { beforeAll, describe, expect, it } from 'vitest';
import { decodeGradientTransport, encodeGradientTransport, parseLinearGradient } from '../utils.js';

// parseColor resolves colours through a canvas; jsdom has none, so hand the
// written colour back as the browser would for these already-normal forms.
beforeAll(() => {
  let fillStyle = '';
  HTMLCanvasElement.prototype.getContext = () => ({
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value) {
      fillStyle = value;
    },
  });
});

describe('parseLinearGradient', () => {
  it('reads an angle and spreads missing stop positions evenly', () => {
    expect(parseLinearGradient('linear-gradient(90deg, rgb(15, 118, 110), rgb(245, 158, 11))')).toEqual({
      angle: 90,
      stops: [
        { hex: '0F766E', opacity: 1, pos: 0 },
        { hex: 'F59E0B', opacity: 1, pos: 1 },
      ],
    });
    const three = parseLinearGradient('linear-gradient(#fff, #888, #000 80%)');
    expect(three.angle).toBe(180);
    expect(three.stops.map((stop) => stop.pos)).toEqual([0, 0.4, 0.8]);
  });

  it('maps side keywords and keeps stop transparency', () => {
    const gradient = parseLinearGradient('linear-gradient(to right, rgba(0, 0, 0, 0.5) 10%, rgb(255, 255, 255) 90%)');
    expect(gradient.angle).toBe(90);
    expect(gradient.stops[0]).toMatchObject({ hex: '000000', opacity: 0.5, pos: 0.1 });
    expect(parseLinearGradient('linear-gradient(to top left, #000, #fff)').angle).toBe(315);
  });

  it('declines what a native fill cannot hold', () => {
    expect(parseLinearGradient('radial-gradient(#000, #fff)')).toBeNull();
    expect(parseLinearGradient('linear-gradient(#000, #fff), linear-gradient(#fff, #000)')).toBeNull();
    expect(parseLinearGradient('linear-gradient(#000 20px, #fff 80px)')).toBeNull();
    expect(parseLinearGradient('linear-gradient(#000)')).toBeNull();
  });

  it('round-trips through the shape-name transport', () => {
    const gradient = parseLinearGradient('linear-gradient(45deg, rgba(15, 118, 110, 0.25), #f59e0b)');
    const payload = encodeGradientTransport(gradient);
    expect(payload).toMatch(/^[A-Za-z0-9_-]+$/);
    expect(decodeGradientTransport(payload)).toEqual(gradient);
  });
});
