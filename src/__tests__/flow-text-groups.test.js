import { beforeEach, describe, expect, it, vi } from 'vitest';
import { exportToPptx } from '../index.js';
import { collectTextParts, isTextContainer } from '../utils.js';

const mockAddText = vi.fn();
const mockAddSlide = vi.fn(() => ({
  addText: mockAddText,
  addShape: vi.fn(),
  addImage: vi.fn(),
  addTable: vi.fn(),
}));

vi.mock('pptxgenjs', () => ({
  default: vi.fn().mockImplementation(function () {
    return {
      defineLayout: vi.fn(),
      addSlide: mockAddSlide,
      write: vi.fn(() => Promise.resolve('')),
      ShapeType: { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse' },
    };
  }),
}));

function rect(left, top, width, height) {
  return { left, top, width, height, right: left + width, bottom: top + height, x: left, y: top };
}

describe('flow text groups', () => {
  beforeEach(() => {
    mockAddText.mockClear();
    mockAddSlide.mockClear();
    document.body.replaceChildren();
    HTMLCanvasElement.prototype.getContext = () => ({
      get fillStyle() {
        return '#000000';
      },
      set fillStyle(_) {},
    });
  });

  it('keeps an inline title and following paragraph in one parent-width text flow', async () => {
    const slide = document.createElement('section');
    slide.className = 'slide';
    slide.style.cssText = 'position:relative;width:960px;height:540px';
    slide.getBoundingClientRect = () => rect(0, 0, 960, 540);

    const card = document.createElement('div');
    card.style.cssText = [
      'position:absolute',
      'left:100px',
      'top:80px',
      'width:320px',
      'height:180px',
      'padding:20px',
      'background:#eef4ee',
    ].join(';');
    card.innerHTML =
      '<b>Verbindlicher Rahmen</b><p>Gemeinsame Ergebnisse, Kapazität und Entscheidungen werden zur Voraussetzung.</p>';
    card.getBoundingClientRect = () => rect(100, 80, 360, 220);

    const paragraph = card.querySelector('p');
    paragraph.style.cssText = 'display:block;margin:12px 0 0';
    paragraph.getBoundingClientRect = () => rect(120, 146, 320, 64);
    const title = card.querySelector('b');
    title.style.cssText = 'display:inline;font-weight:700';
    title.getBoundingClientRect = () => rect(120, 100, 150, 24);

    slide.append(card);
    document.body.append(slide);

    await exportToPptx(slide, { skipDownload: true, skipNormalize: true });

    const flowCalls = mockAddText.mock.calls.filter(([parts]) => Array.isArray(parts));
    expect(flowCalls).toHaveLength(1);

    const [parts, options] = flowCalls[0];
    expect(parts.map((part) => part.text).join('')).toContain('Verbindlicher Rahmen');
    expect(parts.map((part) => part.text).join('')).toContain('Gemeinsame Ergebnisse');
    expect(parts.some((part) => part.options?.breakLine)).toBe(true);
    expect(parts.find((part) => part.text.includes('Verbindlicher Rahmen')).options.bold).toBe(true);
    expect(parts.find((part) => part.text.includes('Gemeinsame Ergebnisse')).options.paraSpaceBefore).toBeCloseTo(9, 4);
    expect(options.w).toBeCloseTo(3.75, 4);
    expect(options.h).toBeCloseTo(2.2917, 4);
  });

  it('keeps an inline lead and a block-displayed semantic text run in one flow', () => {
    const container = document.createElement('div');
    container.innerHTML =
      '<strong>Nur das Zusammenspiel schafft Wirkung.</strong>' +
      '<span style="display:block;margin-top:7px">Ziel, Mittel und Verantwortung greifen ineinander.</span>';
    document.body.append(container);

    expect(isTextContainer(container)).toBe(true);

    const parts = collectTextParts(container, window.getComputedStyle(container), 1);
    expect(parts.map((part) => part.text).join('')).toContain('Nur das Zusammenspiel schafft Wirkung.');
    expect(parts.map((part) => part.text).join('')).toContain(
      'Ziel, Mittel und Verantwortung greifen ineinander.'
    );
    const secondRun = parts.findIndex((part) => part.text.includes('Ziel, Mittel'));
    expect(parts.slice(0, secondRun).some((part) => part.options?.breakLine)).toBe(true);
  });

  it('does not collapse a decorated or positioned child into text flow', () => {
    const container = document.createElement('div');
    container.innerHTML = '<b>Lead</b><p>Body</p>';
    const paragraph = container.querySelector('p');
    paragraph.style.cssText = 'display:block;position:absolute';

    document.body.append(container);

    expect(isTextContainer(container)).toBe(false);
  });

  it('does not collapse div-based metric values and labels into one text flow', () => {
    const metric = document.createElement('div');
    metric.innerHTML =
      '<div class="value">1,8</div><div class="unit">Mio.</div><div class="label">Vorgänge pro Jahr</div>';
    document.body.append(metric);

    expect(isTextContainer(metric)).toBe(false);
  });

  it('does not collapse a flow whose direct children disagree on horizontal alignment', () => {
    const container = document.createElement('div');
    container.style.textAlign = 'left';
    container.innerHTML = '<b style="text-align:left">Titel</b><p style="text-align:right">Körper</p>';
    document.body.append(container);

    expect(isTextContainer(container)).toBe(false);
  });

  it('does not collapse a flow whose direct children disagree on wrapping policy', () => {
    const container = document.createElement('div');
    container.innerHTML = '<b>Titel</b><p style="white-space:nowrap">Diese Zeile darf im Browser nicht umbrechen.</p>';
    document.body.append(container);

    expect(isTextContainer(container)).toBe(false);
  });

  it('keeps consecutive heading blocks as distinct PPTX paragraphs', () => {
    const container = document.createElement('div');
    container.innerHTML = '<h3>Leitgedanke</h3><h4>Konsequenz</h4><p>Die Umsetzung folgt der Entscheidung.</p>';
    document.body.append(container);

    const parts = collectTextParts(container, window.getComputedStyle(container), 1);
    const texts = parts.filter((part) => part.text).map((part) => part.text);
    expect(texts).toEqual(['Leitgedanke', 'Konsequenz', 'Die Umsetzung folgt der Entscheidung.']);

    for (let index = 0; index < texts.length - 1; index++) {
      const current = parts.findIndex((part) => part.text === texts[index]);
      const next = parts.findIndex((part, partIndex) => partIndex > current && part.text === texts[index + 1]);
      expect(parts.slice(current + 1, next).some((part) => part.options?.breakLine)).toBe(true);
    }
  });

  it('keeps list structures outside a text-flow group', () => {
    const container = document.createElement('div');
    container.innerHTML = '<b>Entscheidung</b><ul><li>Erster Punkt</li><li>Zweiter Punkt</li></ul>';
    document.body.append(container);

    expect(isTextContainer(container)).toBe(false);
  });

  it.each([
    ['Hello<strong> world</strong>', 'Hello world'],
    ['<strong>Hello </strong>world', 'Hello world'],
    ['A <em>B</em> C', 'A B C'],
  ])('preserves collapsible whitespace across inline run boundaries', (html, expected) => {
    const container = document.createElement('div');
    container.innerHTML = html;
    document.body.append(container);

    const parts = collectTextParts(container, window.getComputedStyle(container), 1);
    expect(parts.map((part) => part.text).join('')).toBe(expected);
  });

  it('keeps an inline-block with its own layout out of a rich-text flow', () => {
    const container = document.createElement('div');
    container.innerHTML =
      '<span style="display:inline-block;width:300px;text-align:right">Right</span><span>After</span>';
    document.body.append(container);

    expect(isTextContainer(container)).toBe(false);
  });

  it('uses computed display rather than the tag name for paragraph boundaries', () => {
    const inlineHeadings = document.createElement('div');
    inlineHeadings.innerHTML = '<h3 style="display:inline">Alpha</h3><h4 style="display:inline">Beta</h4>';
    document.body.append(inlineHeadings);

    const inlineParts = collectTextParts(inlineHeadings, window.getComputedStyle(inlineHeadings), 1);
    expect(inlineParts.map((part) => part.text).join('')).toBe('AlphaBeta');
    expect(inlineParts.some((part) => part.options?.breakLine)).toBe(false);

    const blockHeadings = document.createElement('div');
    blockHeadings.innerHTML = '<h3>Alpha</h3><h4>Beta</h4>';
    document.body.append(blockHeadings);

    const blockParts = collectTextParts(blockHeadings, window.getComputedStyle(blockHeadings), 1);
    expect(blockParts.some((part) => part.options?.breakLine)).toBe(true);
  });
});
