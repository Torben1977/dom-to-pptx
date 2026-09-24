import { describe, it, expect, beforeAll, vi } from 'vitest';
import { extractTableData } from '../utils.js';
import { exportToPptx } from '../index.js';

// Mock pptxgenjs
const mockAddText = vi.fn();
const mockAddSlide = vi.fn(() => ({
  addText: mockAddText,
  addShape: vi.fn(),
  addImage: vi.fn(),
  addTable: vi.fn(),
}));

vi.mock('pptxgenjs', () => {
  return {
    default: vi.fn().mockImplementation(function () {
      return {
        defineLayout: vi.fn(),
        addSlide: mockAddSlide,
        write: vi.fn(() => Promise.resolve('')),
      };
    }),
  };
});

describe('extractTableData', () => {
  beforeAll(() => {
    // Mock HTMLCanvasElement.prototype.getContext for JSDOM env
    let fillStyle = '';
    HTMLCanvasElement.prototype.getContext = () => ({
      get fillStyle() {
        return fillStyle;
      },
      set fillStyle(val) {
        fillStyle = val;
      },
      clearRect: () => {},
      fillRect: () => {},
      getImageData: () => ({ data: [0, 0, 0, 0] }),
    });
  });

  it('extracts table rows and columns correctly', () => {
    const table = document.createElement('table');
    table.innerHTML = `
      <tr>
        <td>Cell 1</td>
        <td>Cell 2</td>
      </tr>
    `;
    document.body.appendChild(table);

    const data = extractTableData(table, 1);
    expect(data.rows.length).toBe(1);
    expect(data.rows[0].length).toBe(2);
    expect(data.rows[0][0].text[0].text).toBe('Cell 1');

    document.body.removeChild(table);
  });

  // `border-spacing` has no effect once the borders are collapsed, but the
  // browser keeps reporting its value. Adding it anyway gave every collapsed
  // table cell margins wider than its CSS padding, so its cells wrapped earlier
  // than in the browser — which pushed the rows below down and cost words.
  it('keeps a collapsed table cell margin at its CSS padding', () => {
    const table = document.createElement('table');
    table.setAttribute('style', 'border-collapse:collapse;border-spacing:2px');
    table.innerHTML = '<tr><td style="padding:8px 10px">Dimension</td></tr>';
    document.body.appendChild(table);

    try {
      // PptxGenJS takes cell margins as [top, right, bottom, left] in points.
      const margin = extractTableData(table, 1).rows[0][0].options.margin;
      expect(margin.map((value) => Number(value.toFixed(4)))).toEqual([6, 7.5, 6, 7.5]);
    } finally {
      table.remove();
    }
  });

  it('adds border-spacing to the cell margin only where the browser applies it', () => {
    const table = document.createElement('table');
    table.setAttribute('style', 'border-collapse:separate;border-spacing:2px');
    table.innerHTML = '<tr><td style="padding:8px 10px">Dimension</td></tr>';
    document.body.appendChild(table);

    try {
      // Half of the 2px gap belongs to each of the two cells that share it.
      const margin = extractTableData(table, 1).rows[0][0].options.margin;
      expect(margin.map((value) => Number(value.toFixed(4)))).toEqual([6.75, 8.25, 6.75, 8.25]);
    } finally {
      table.remove();
    }
  });

  // The measured height travels as a row minimum. It is not what fixed the
  // table probe — the cell margins were — and the fidelity fixture has no row
  // the browser makes taller than its text, so this guards the transmission
  // only, not the rendering.
  it('carries the measured row height alongside the rows it belongs to', () => {
    const table = document.createElement('table');
    table.innerHTML = '<tr><td>Kopf</td></tr><tr></tr><tr><td>Zeile</td></tr>';
    document.body.appendChild(table);
    const rows = Array.from(table.querySelectorAll('tr'));
    rows[0].getBoundingClientRect = () => ({ height: 48, width: 200, top: 0, left: 0, right: 200, bottom: 48 });
    rows[1].getBoundingClientRect = () => ({ height: 999, width: 200, top: 48, left: 0, right: 200, bottom: 1047 });
    rows[2].getBoundingClientRect = () => ({ height: 96, width: 200, top: 48, left: 0, right: 200, bottom: 144 });

    try {
      const data = extractTableData(table, 1);
      // The empty row never becomes a row, so its height must not become one either.
      expect(data.rows).toHaveLength(2);
      expect(data.rowHeights).toEqual([48 / 96, 96 / 96]);
    } finally {
      table.remove();
    }
  });

  it('maps writing-mode to textDirection in table cells', () => {
    const table = document.createElement('table');
    table.innerHTML = `
      <tr>
        <td style="writing-mode: vertical-rl;">Vertical Cell</td>
        <td style="writing-mode: vertical-lr; text-orientation: upright;">Upright Cell</td>
        <td>Normal Cell</td>
      </tr>
    `;
    document.body.appendChild(table);

    const data = extractTableData(table, 1);
    expect(data.rows[0][0].options.textDirection).toBe('vert');
    expect(data.rows[0][1].options.textDirection).toBe('wordArtVert');
    expect(data.rows[0][2].options.textDirection).toBeUndefined();

    document.body.removeChild(table);
  });

  it('flattens translucent table backgrounds and borders against parent backdrop', () => {
    // Create a container with dark background (representing the slide)
    const container = document.createElement('div');
    container.style.backgroundColor = '#101018'; // opaque parent backdrop
    document.body.appendChild(container);

    const table = document.createElement('table');
    table.innerHTML = `
      <tr>
        <td style="background-color: rgba(255, 255, 255, 0.04); border-top-style: solid; border-top-width: 1px; border-top-color: rgba(255, 255, 255, 0.08); color: #fff;">
          Translucent Cell
        </td>
      </tr>
    `;
    container.appendChild(table);

    const data = extractTableData(table, 1);
    expect(data.rows.length).toBe(1);
    expect(data.rows[0].length).toBe(1);

    const cellOptions = data.rows[0][0].options;
    // Expected background: rgba(255,255,255,0.04) blended over #101018 (which is r=16, g=16, b=24)
    // r = 255 * 0.04 + 16 * 0.96 = 10.2 + 15.36 = 25.56 => 26 (0x1a)
    // g = 255 * 0.04 + 16 * 0.96 = 25.56 => 26 (0x1a)
    // b = 255 * 0.04 + 24 * 0.96 = 10.2 + 23.04 = 33.24 => 33 (0x21)
    // => Hex: '1A1A21'
    expect(cellOptions.fill).toEqual({ color: '1A1A21' });

    // Expected top border: rgba(255,255,255,0.08) blended over its own cell background (which is #1A1A21)
    // r = 255 * 0.08 + 26 * 0.92 = 20.4 + 23.92 = 44.32 => 44 (0x2c)
    // g = 255 * 0.08 + 26 * 0.92 = 44 (0x2c)
    // b = 255 * 0.08 + 33 * 0.92 = 20.4 + 30.36 = 50.76 => 51 (0x33)
    // => Hex: '2C2C33'
    expect(cellOptions.border[0].color).toBe('2C2C33');

    document.body.removeChild(container);
  });

  it('correctly calculates and sets bullet.indent for list items with padding-left', async () => {
    mockAddText.mockClear();

    // Create list in the DOM
    const container = document.createElement('div');
    container.className = 'slide';
    container.style.width = '960px';
    container.style.height = '540px';

    const ul = document.createElement('ul');
    ul.style.paddingLeft = '20px';

    const li = document.createElement('li');
    li.style.paddingLeft = '100px';
    li.textContent = 'Indented text';

    ul.appendChild(li);
    container.appendChild(ul);
    document.body.appendChild(container);

    // Mock getBoundingClientRect for JSDOM layout engine
    container.getBoundingClientRect = () => ({ width: 960, height: 540, left: 0, top: 0, right: 960, bottom: 540 });
    ul.getBoundingClientRect = () => ({ width: 900, height: 400, left: 20, top: 20, right: 920, bottom: 420 });
    li.getBoundingClientRect = () => ({ width: 800, height: 50, left: 20, top: 20, right: 820, bottom: 70 });

    // Call exportToPptx
    await exportToPptx(container, { skipDownload: true, skipNormalize: true });

    // Assert mockAddText was called with the correct bullet.indent
    expect(mockAddText).toHaveBeenCalled();

    let foundRun = null;
    for (const call of mockAddText.mock.calls) {
      const [textParts] = call;
      if (Array.isArray(textParts)) {
        const run = textParts.find((part) => part.text === 'Indented text');
        if (run) {
          foundRun = run;
          break;
        }
      }
    }

    expect(foundRun).not.toBeNull();
    expect(foundRun.options.bullet).not.toBeNull();
    // Bullet indent gap stays at standard 20pt, while hierarchical indentLevel is set for sub-bullets
    expect(foundRun.options.bullet.indent).toBe(20);
    // extraIndentPx = 100px -> indentLevel = min(8, round(100 / 20)) = 5
    expect(foundRun.options.indentLevel).toBe(5);

    document.body.removeChild(container);
  });

  it('falls back to default 20pt bullet.indent when calculated indent is 0', async () => {
    mockAddText.mockClear();

    // Create list in the DOM with no padding-left overrides
    const container = document.createElement('div');
    container.className = 'slide';
    container.style.width = '960px';
    container.style.height = '540px';

    const ul = document.createElement('ul');

    const li = document.createElement('li');
    li.textContent = 'Normal text';

    ul.appendChild(li);
    container.appendChild(ul);
    document.body.appendChild(container);

    // Mock getBoundingClientRect for JSDOM layout engine
    container.getBoundingClientRect = () => ({ width: 960, height: 540, left: 0, top: 0, right: 960, bottom: 540 });
    ul.getBoundingClientRect = () => ({ width: 900, height: 400, left: 0, top: 0, right: 900, bottom: 400 });
    li.getBoundingClientRect = () => ({ width: 900, height: 50, left: 0, top: 0, right: 900, bottom: 50 });

    // Call exportToPptx
    await exportToPptx(container, { skipDownload: true, skipNormalize: true });

    // Assert mockAddText was called with the default bullet.indent
    expect(mockAddText).toHaveBeenCalled();

    let foundRun = null;
    for (const call of mockAddText.mock.calls) {
      const [textParts] = call;
      if (Array.isArray(textParts)) {
        const run = textParts.find((part) => part.text === 'Normal text');
        if (run) {
          foundRun = run;
          break;
        }
      }
    }

    expect(foundRun).not.toBeNull();
    expect(foundRun.options.bullet).not.toBeNull();
    expect(foundRun.options.bullet.indent).toBe(20); // Default fallback

    document.body.removeChild(container);
  });
});
