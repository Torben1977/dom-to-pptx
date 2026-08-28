import { describe, expect, it } from 'vitest';
import { generateGradientSVG } from '../utils.js';

function gradientEndpoints(direction) {
  const dataUrl = generateGradientSVG(400, 200, `linear-gradient(${direction}, #ff0000, #0000ff)`, 0);
  expect(dataUrl).toMatch(/^data:image\/svg\+xml;base64,/);
  const svg = atob(dataUrl.slice(dataUrl.indexOf(',') + 1));
  const documentNode = new DOMParser().parseFromString(svg, 'image/svg+xml');
  const gradient = documentNode.getElementsByTagName('linearGradient')[0];
  expect(gradient).toBeDefined();
  return {
    x1: gradient.getAttribute('x1'),
    y1: gradient.getAttribute('y1'),
    x2: gradient.getAttribute('x2'),
    y2: gradient.getAttribute('y2'),
  };
}

describe('CSS gradient angle semantics', () => {
  it('maps 0deg and 90deg to the same axes as their CSS direction keywords', () => {
    expect(gradientEndpoints('0deg')).toEqual(gradientEndpoints('to top'));
    expect(gradientEndpoints('90deg')).toEqual(gradientEndpoints('to right'));
  });

  it('converts turn units before calculating the gradient vector', () => {
    expect(gradientEndpoints('0.25turn')).toEqual(gradientEndpoints('90deg'));
    expect(gradientEndpoints('-0.25turn')).toEqual(gradientEndpoints('-90deg'));
  });

  it('converts grad units before calculating the gradient vector', () => {
    expect(gradientEndpoints('100grad')).toEqual(gradientEndpoints('90deg'));
    expect(gradientEndpoints('-100grad')).toEqual(gradientEndpoints('-90deg'));
  });

  it('projects diagonal angles across a non-square box instead of a unit square', () => {
    expect(gradientEndpoints('45deg')).toEqual({
      x1: '12.5%',
      y1: '125%',
      x2: '87.5%',
      y2: '-25%',
    });
  });
});
