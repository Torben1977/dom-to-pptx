import { describe, expect, it } from 'vitest';
import { isOutOfTextFlow } from '../utils.js';

// One question decides whether content may be folded into a shared PowerPoint
// text run: does the browser paint it where the text flow puts it? A run
// carries characters, never a placement of its own.

const style = (overrides) => ({
  position: 'static',
  float: 'none',
  transform: 'none',
  top: 'auto',
  right: 'auto',
  bottom: 'auto',
  left: 'auto',
  ...overrides,
});

describe('isOutOfTextFlow', () => {
  it('keeps ordinary in-flow content in the text flow', () => {
    expect(isOutOfTextFlow(style({}))).toBe(false);
    expect(isOutOfTextFlow(style({ display: 'block' }))).toBe(false);
  });

  it('treats absolutely positioned and fixed content as painted elsewhere', () => {
    expect(isOutOfTextFlow(style({ position: 'absolute', left: '0px' }))).toBe(true);
    expect(isOutOfTextFlow(style({ position: 'fixed' }))).toBe(true);
  });

  it('treats floated content as painted elsewhere', () => {
    expect(isOutOfTextFlow(style({ float: 'left' }))).toBe(true);
    expect(isOutOfTextFlow(style({ float: 'right' }))).toBe(true);
  });

  it('treats transformed content as painted elsewhere', () => {
    expect(isOutOfTextFlow(style({ transform: 'translateX(10px)' }))).toBe(true);
    expect(isOutOfTextFlow(style({ transform: 'rotate(2deg)' }))).toBe(true);
    expect(isOutOfTextFlow(style({ transform: 'matrix(1, 0, 0, 1, 12, 0)' }))).toBe(true);
  });

  // A painting hint such as `translateZ(0)` resolves to the identity matrix and
  // moves nothing. Reading it as out-of-flow cost a whole table its editable
  // cells, because content needing its own box inside a cell rasterizes the table.
  it('keeps content with an identity transform in the text flow', () => {
    expect(isOutOfTextFlow(style({ transform: 'matrix(1, 0, 0, 1, 0, 0)' }))).toBe(false);
    expect(isOutOfTextFlow(style({ transform: 'matrix3d(1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1)' }))).toBe(
      false
    );
  });

  // `relative` and `sticky` stay in flow even when offset. Rejecting them would
  // give each shifted span a text box of its own, and those boxes overlap as
  // soon as Office wraps differently — measured at five overlaps on the
  // relative-offset probe against none when the span stays in the run. A word
  // painted a few points off is the smaller error, so this is deliberate.
  it('keeps relative and sticky content in the text flow, offset or not', () => {
    expect(isOutOfTextFlow(style({ position: 'relative' }))).toBe(false);
    expect(isOutOfTextFlow(style({ position: 'relative', top: '-10px' }))).toBe(false);
    expect(isOutOfTextFlow(style({ position: 'sticky', top: '10px' }))).toBe(false);
  });

  it('survives a missing or partial style object', () => {
    expect(isOutOfTextFlow(null)).toBe(false);
    expect(isOutOfTextFlow({})).toBe(false);
    expect(isOutOfTextFlow({ position: 'absolute' })).toBe(true);
  });
});
