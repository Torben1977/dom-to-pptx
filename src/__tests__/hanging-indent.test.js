import { afterEach, describe, expect, it } from 'vitest';
import {
  applyHangingIndent,
  HANGING_INDENT_BULLET_CODE,
  measureMarkerHangPx,
  resolveHangingIndentPt,
} from '../utils.js';

// A hanging first line and a list marker are the same problem: the first line
// starts left of the ones below it. OpenXML has one way to say that, the pair
// marL/indent, and the room it needs has to come out of the box's left inset
// rather than being added on top of it.

const mount = (html) => {
  const host = document.createElement('div');
  host.innerHTML = html;
  document.body.appendChild(host);
  return host.firstElementChild;
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('resolveHangingIndentPt', () => {
  it('turns a negative text-indent into a hanging indent in points', () => {
    const node = mount('<p style="padding-left:28px;text-indent:-28px">Erster Punkt</p>');
    expect(resolveHangingIndentPt(node, window.getComputedStyle(node), 1)).toBeCloseTo(21, 5);
    expect(resolveHangingIndentPt(node, window.getComputedStyle(node), 0.5)).toBeCloseTo(10.5, 5);
  });

  // The browser lets the first line spill out of the box to the left; a
  // PowerPoint text frame has nothing left of its inset.
  it('stops the hang at the left padding', () => {
    const node = mount('<p style="padding-left:10px;text-indent:-28px">Erster Punkt</p>');
    expect(resolveHangingIndentPt(node, window.getComputedStyle(node), 1)).toBeCloseTo(7.5, 5);
  });

  it('ignores an indent that has no OpenXML counterpart', () => {
    const positive = mount('<p style="padding-left:28px;text-indent:28px">Erster Punkt</p>');
    expect(resolveHangingIndentPt(positive, window.getComputedStyle(positive), 1)).toBe(0);

    const relative = mount('<p style="padding-left:28px;text-indent:-10%">Erster Punkt</p>');
    expect(resolveHangingIndentPt(relative, window.getComputedStyle(relative), 1)).toBe(0);

    const none = mount('<p style="padding-left:28px">Erster Punkt</p>');
    expect(resolveHangingIndentPt(none, window.getComputedStyle(none), 1)).toBe(0);
  });

  // text-indent is inherited but takes effect per block box, while the shape has
  // one inset for all of its paragraphs.
  it('steps back when a block inside the shape wants a different indent', () => {
    const node = mount(
      '<div style="padding-left:28px;text-indent:-28px">' +
        '<p>Erster Punkt</p><p style="text-indent:0">Zweiter Punkt</p></div>'
    );
    expect(resolveHangingIndentPt(node, window.getComputedStyle(node), 1)).toBe(0);
  });

  it('is not disturbed by inline content', () => {
    const node = mount(
      '<p style="padding-left:28px;text-indent:-28px">Erster <strong>Punkt</strong> mit <em>Auszeichnung</em></p>'
    );
    expect(resolveHangingIndentPt(node, window.getComputedStyle(node), 1)).toBeCloseTo(21, 5);
  });
});

describe('applyHangingIndent', () => {
  const sentinel = { characterCode: HANGING_INDENT_BULLET_CODE, indent: 21 };

  it('opens every paragraph with the indent and leaves the runs inside it alone', () => {
    const parts = [
      { text: 'Erster ' },
      { text: 'Punkt', options: { bold: true, breakLine: true } },
      { text: 'Zweiter Punkt' },
    ];

    applyHangingIndent(parts, 21);

    expect(parts.map((part) => part.options.bullet)).toEqual([sentinel, undefined, sentinel]);
  });

  it('never displaces a bullet the list path already decided on', () => {
    const listBullet = { code: '2022', indent: 20 };
    const parts = [{ text: 'Punkt', options: { bullet: listBullet } }];

    applyHangingIndent(parts, 21);

    expect(parts[0].options.bullet).toBe(listBullet);
  });
});

describe('measureMarkerHangPx', () => {
  const styleOf = (declaration) => {
    const node = mount(`<ul style="${declaration}"><li>Punkt</li></ul>`);
    return window.getComputedStyle(node.firstElementChild);
  };

  it('reports nothing for an item that has no marker', () => {
    expect(measureMarkerHangPx('ul', styleOf('list-style-type:none'), null)).toBe(0);
  });

  // jsdom has no layout, so the probe measures zero. Callers read that as
  // "keep doing what you did before" rather than as a distance of zero.
  it('reports nothing instead of throwing where nothing can be measured', () => {
    expect(measureMarkerHangPx('ul', styleOf('list-style-type:disc;font-size:17px'), null)).toBe(0);
    expect(measureMarkerHangPx('ol', styleOf('list-style-type:decimal;font-size:18px'), null, 7)).toBe(0);
  });

  it('leaves no probe behind in the document', () => {
    measureMarkerHangPx('ul', styleOf('list-style-type:square;font-size:19px'), null);
    expect(document.querySelectorAll('[data-pptx-marker-probe]')).toHaveLength(0);
  });
});
