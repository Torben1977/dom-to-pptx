// src/__tests__/node-exporter.test.js
//
// Tests for the Puppeteer launch-arg helper used by exportHtmlToPptx.
//
// Motivation: without --allow-file-access-from-files, Chromium blocks
// fetch() from a file:// page to other file:// URLs — which is the
// exact code path autoEmbedFonts uses to pull local .ttf files into an
// embedded font blob. The failure is silent (fetch just rejects, the
// per-font warn is easy to miss), so users end up with a fonts-less
// PPTX that PowerPoint renders in a fallback system font. Enabling the
// flag by default prevents that failure mode.

import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import path from 'node:path';
import { getLaunchArgs, exportHtmlToPptx } from '../node-exporter.js';

describe('getLaunchArgs', () => {
  it('enables --allow-file-access-from-files for Chrome/Chromium/Edge', () => {
    const args = getLaunchArgs('chrome');
    expect(args).toContain('--allow-file-access-from-files');
  });

  it('keeps the existing --no-sandbox flags for Chrome-family browsers', () => {
    const args = getLaunchArgs('chrome');
    expect(args).toContain('--no-sandbox');
    expect(args).toContain('--disable-setuid-sandbox');
  });

  it('returns an empty arg list for Firefox (WebDriver BiDi manages its own launch)', () => {
    expect(getLaunchArgs('firefox')).toEqual([]);
  });

  it('is a pure function — repeated calls produce distinct arrays that can be safely mutated', () => {
    const a = getLaunchArgs('chrome');
    const b = getLaunchArgs('chrome');
    expect(a).toEqual(b);
    // Mutating one should not affect the other.
    a.push('--custom');
    expect(b).not.toContain('--custom');
  });

  it('exports slides with custom dimensions correctly', async () => {
    const html = `
      <!doctype html>
      <html>
      <body style="margin:0">
        <div class="slide" style="width:1920px;height:1080px;position:relative;background:#fff">
          <h1>Slide One</h1>
        </div>
      </body>
      </html>
    `;
    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: {
        width: 13.333333,
        height: 7.5,
      },
    });

    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/presentation.xml').async('string');

    const sldSzMatch = xml.match(/<[a-zA-Z0-9:]*sldSz\s+([^>]+)>/);
    expect(sldSzMatch).not.toBeNull();
    const attrs = sldSzMatch[1];
    const cxMatch = attrs.match(/cx=["'](\d+)["']/);
    const cyMatch = attrs.match(/cy=["'](\d+)["']/);
    expect(cxMatch).not.toBeNull();
    expect(cyMatch).not.toBeNull();

    const cx = parseInt(cxMatch[1], 10);
    const cy = parseInt(cyMatch[1], 10);

    const width = cx / 914400;
    const height = cy / 914400;

    expect(width).toBeCloseTo(13.333333, 4);
    expect(height).toBeCloseTo(7.5, 4);
  }, 30000);

  it('exports pseudo-elements (linear-gradients and border triangles) correctly', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
      body { margin: 0; }
      .slide {
        width: 1920px;
        height: 1080px;
        position: relative;
        background: #fff;
      }
      .card {
        position: absolute;
        left: 200px;
        top: 200px;
        width: 600px;
        height: 300px;
      }
      .card::before {
        content: "";
        position: absolute;
        left: 0;
        right: 0;
        top: 0;
        height: 12px;
        background: linear-gradient(90deg, #00479d, #1871c5);
      }
      .arrow {
        position: absolute;
        left: 950px;
        top: 350px;
        width: 180px;
        height: 6px;
      }
      .arrow::after {
        content: "";
        position: absolute;
        right: -1px;
        top: 50%;
        transform: translateY(-50%);
        width: 0;
        height: 0;
        border-left: 30px solid #ed6f18;
        border-top: 20px solid transparent;
        border-bottom: 20px solid transparent;
      }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="card"></div>
          <div class="arrow"></div>
        </div>
      </body>
      </html>
    `;
    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: {
        includePseudoElements: true,
      },
    });

    const zip = await JSZip.loadAsync(buffer);
    const mediaFiles = Object.keys(zip.files).filter((k) => k.startsWith('ppt/media/'));
    expect(mediaFiles.length).toBeGreaterThanOrEqual(2);
  }, 40000);

  it('exports an inline lead and following paragraph as one parent-width text flow', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
        body { margin: 0; }
        .slide { width: 1920px; height: 1080px; position: relative; background: #fff; }
        .card { position: absolute; left: 180px; top: 180px; width: 640px; padding: 40px; background: #eef4ee; }
        .card b { display: inline; font: 700 42px/1.15 Arial; }
        .card p { display: block; margin: 20px 0 0; font: 400 30px/1.35 Arial; }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="card"><b>Verbindlicher Rahmen</b><p>Gemeinsame Ergebnisse, Kapazität und Entscheidungen werden zur Voraussetzung.</p></div>
        </div>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5 },
    });

    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const titleIndex = xml.indexOf('<a:t>Verbindlicher Rahmen</a:t>');
    const bodyIndex = xml.indexOf(
      '<a:t>Gemeinsame Ergebnisse, Kapazität und Entscheidungen werden zur Voraussetzung.</a:t>'
    );
    expect(titleIndex).toBeGreaterThan(-1);
    expect(bodyIndex).toBeGreaterThan(titleIndex);

    const shapeStart = xml.lastIndexOf('<p:sp>', titleIndex);
    const shapeEnd = xml.indexOf('</p:sp>', titleIndex);
    const textShape = xml.slice(shapeStart, shapeEnd);
    expect(textShape).toContain(
      '<a:t>Gemeinsame Ergebnisse, Kapazität und Entscheidungen werden zur Voraussetzung.</a:t>'
    );

    const widthMatch = textShape.match(/<a:ext cx="(\d+)"/);
    expect(widthMatch).not.toBeNull();
    expect(Number(widthMatch[1])).toBeGreaterThan(3_000_000);
  }, 40000);

  it('uses PowerPoint autofit only for browser-overflowing text boxes', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
        body { margin: 0; }
        .slide { width: 1920px; height: 1080px; position: relative; background: #fff; }
        .safe, .near-boundary, .overflowing {
          position: absolute;
          left: 180px;
          overflow: hidden;
        }
        .safe {
          top: 180px;
          width: 300px;
          height: 130px;
          padding: 27px 25px;
          border-radius: 70px;
          font: 700 22px/29px Arial;
          text-align: center;
        }
        .overflowing {
          top: 480px;
          width: 420px;
          height: 52px;
          padding: 24px;
          font: 700 32px/40px Arial;
        }
        .near-boundary {
          top: 360px;
          width: 300px;
          height: 100px;
          font: 700 84px/100px Arial;
          white-space: nowrap;
        }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="safe">Browser-fit title<br>keeps its size<br><span style="font-size:16px;font-weight:400">with safe reserve</span></div>
          <div class="near-boundary">Fit</div>
          <div class="overflowing">Browser-overflowing title with a second line</div>
        </div>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5 },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const shapeFor = (text) => {
      const textIndex = xml.indexOf(`<a:t>${text}</a:t>`);
      expect(textIndex).toBeGreaterThan(-1);
      return xml.slice(xml.lastIndexOf('<p:sp>', textIndex), xml.indexOf('</p:sp>', textIndex));
    };

    expect(shapeFor('Browser-fit title')).not.toContain('<a:normAutofit');
    expect(shapeFor('Fit')).not.toContain('<a:normAutofit');
    expect(shapeFor('Browser-overflowing title with a second line')).toContain('<a:normAutofit');
  }, 40000);

  it('keeps the representative OrgLith deck.html text-flow patterns together', async () => {
    const fixture = path.resolve('src/__tests__/fixtures/orglith-text-flow-deck.html');
    const buffer = await exportHtmlToPptx(fixture, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5 },
    });
    const zip = await JSZip.loadAsync(buffer);

    const cases = [
      [1, 'Gemeinsam messen', 'Durchlaufzeit und Rückfragen als Ergebnisgrößen.'],
      [2, 'Weiter wie bisher', 'Lokale Optimierung, Verantwortung und Ressourcen bleiben verteilt.'],
      [2, 'Verbindlicher Rahmen', 'Gemeinsame Ergebnisse, Kapazität und Entscheidungen werden zur Voraussetzung.'],
      [3, 'Einsparungen', 'Es gibt noch keine belastbare Schätzung der Einsparungen.'],
    ];

    for (const [slideNo, title, paragraph] of cases) {
      const xml = await zip.file(`ppt/slides/slide${slideNo}.xml`).async('string');
      const titleIndex = xml.indexOf(`<a:t>${title}</a:t>`);
      expect(titleIndex).toBeGreaterThan(-1);
      const shapeStart = xml.lastIndexOf('<p:sp>', titleIndex);
      const shapeEnd = xml.indexOf('</p:sp>', titleIndex);
      const textShape = xml.slice(shapeStart, shapeEnd);
      expect(textShape).toContain(`<a:t>${paragraph}`);

      const widthMatch = textShape.match(/<a:ext cx="(\d+)"/);
      expect(widthMatch).not.toBeNull();
      expect(Number(widthMatch[1])).toBeGreaterThan(1_000_000);
    }

    const metricXml = await zip.file('ppt/slides/slide4.xml').async('string');
    const shapeFor = (text) => {
      const textIndex = metricXml.indexOf(`<a:t>${text}</a:t>`);
      expect(textIndex).toBeGreaterThan(-1);
      const shapeStart = metricXml.lastIndexOf('<p:sp>', textIndex);
      const shapeEnd = metricXml.indexOf('</p:sp>', textIndex);
      return metricXml.slice(shapeStart, shapeEnd);
    };

    // A metric's nested divs are independent layout objects.  Grouping them as
    // editorial text flow would merge their geometry and reintroduce the
    // value/unit/label overlaps seen in converted decks.
    expect(shapeFor('1,8')).not.toContain('<a:t>Mio.</a:t>');
    expect(shapeFor('1,8')).not.toContain('<a:t>Vorgänge pro Jahr</a:t>');
    expect(shapeFor('2027–')).not.toContain('<a:t>2029</a:t>');

    const boundaryXml = await zip.file('ppt/slides/slide5.xml').async('string');
    const boundaryShapeFor = (text) => {
      const textIndex = boundaryXml.indexOf(`<a:t>${text}</a:t>`);
      expect(textIndex).toBeGreaterThan(-1);
      const shapeStart = boundaryXml.lastIndexOf('<p:sp>', textIndex);
      const shapeEnd = boundaryXml.indexOf('</p:sp>', textIndex);
      return boundaryXml.slice(shapeStart, shapeEnd);
    };

    const headingShape = boundaryShapeFor('Leitgedanke');
    expect(headingShape).toContain('<a:t>Konsequenz</a:t>');
    expect(headingShape).toContain('<a:t>Die Umsetzung folgt der Entscheidung.</a:t>');
    expect((headingShape.match(/<a:p>/g) || []).length).toBeGreaterThanOrEqual(3);

    // Per-paragraph alignment cannot be represented by one text shape.  Keep
    // those boxes independent rather than silently inheriting the parent style.
    expect(boundaryShapeFor('Getrennte Ausrichtung')).not.toContain(
      '<a:t>Dieser Absatz ist bewusst rechtsbündig.</a:t>'
    );
  }, 40000);

  it('exports textual pseudo-content exactly once with browser-authored spacing', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
        body { margin: 0; }
        .slide { width: 1920px; height: 1080px; position: relative; background: #fff; }
        .metric, .callout {
          position: absolute;
          left: 180px;
          width: 500px;
          height: 100px;
          font: 700 42px/1.2 Arial;
        }
        .metric { top: 180px; }
        .metric::before { content: "€"; }
        .metric::after { content: "%"; }
        .callout { top: 380px; }
        .callout::before {
          content: "!";
          position: absolute;
          left: -60px;
          top: 0;
          width: 44px;
          height: 44px;
          color: white;
          background: #d8342a;
          border-radius: 50%;
          text-align: center;
        }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="metric">31</div>
          <div class="callout">Positioned warning</div>
        </div>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5 },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const textRuns = Array.from(xml.matchAll(/<a:t>(.*?)<\/a:t>/g), (match) => match[1]);

    expect(textRuns.filter((text) => text.trim() === '€')).toHaveLength(1);
    expect(textRuns.filter((text) => text.trim() === '%')).toHaveLength(1);
    expect(textRuns.filter((text) => text.trim() === '!')).toHaveLength(1);

    const metricIndex = xml.indexOf('<a:t>31</a:t>');
    const metricShape = xml.slice(xml.lastIndexOf('<p:sp>', metricIndex), xml.indexOf('</p:sp>', metricIndex));
    expect(Array.from(metricShape.matchAll(/<a:t>(.*?)<\/a:t>/g), (match) => match[1]).join('')).toBe('€31%');

    const warningIndex = xml.indexOf('<a:t>Positioned warning</a:t>');
    const warningShape = xml.slice(xml.lastIndexOf('<p:sp>', warningIndex), xml.indexOf('</p:sp>', warningIndex));
    expect(warningShape).not.toContain('<a:t>!</a:t>');
  }, 40000);

  it('preserves vertical transformed pseudo text', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
        body { margin: 0; }
        .slide { width: 1920px; height: 1080px; position: relative; background: #fff; }
        .axis {
          position: absolute;
          left: 220px;
          top: 160px;
          width: 640px;
          height: 420px;
        }
        .axis::after {
          content: "Kapazitaetswirkung";
          position: absolute;
          left: -48px;
          top: 36px;
          width: 20px;
          height: 336px;
          color: #14532d;
          font: 700 18px/20px Arial;
          writing-mode: vertical-rl;
          transform: rotate(180deg);
          transform-origin: center;
        }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="axis"></div>
        </div>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, includePseudoElements: true },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    const axisTextIndex = xml.indexOf('<a:t>Kapazitaetswirkung</a:t>');
    expect(axisTextIndex).toBeGreaterThan(-1);
    const axisShape = xml.slice(xml.lastIndexOf('<p:sp>', axisTextIndex), xml.indexOf('</p:sp>', axisTextIndex));
    expect(axisShape).toMatch(/<a:xfrm[^>]+rot="-?10800000"/);
    expect(axisShape).toMatch(/<a:bodyPr[^>]+vert="eaVert"/);
  }, 40000);

  it('resolves scoped CSS counters in textual pseudo-content', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
        body { margin: 0; }
        .slide { width: 1920px; height: 1080px; position: relative; background: #fff; }
        .decisions {
          position: absolute;
          left: 980px;
          top: 180px;
          width: 600px;
          counter-reset: decision;
        }
        .decision {
          counter-increment: decision;
          font: 400 30px/1.3 Arial;
        }
        .decision::before {
          content: counter(decision) ". ";
          font-weight: 700;
        }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="decisions">
            <div class="decision">Rahmen beschliessen</div>
            <div class="decision">Wirkung pruefen</div>
          </div>
        </div>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, includePseudoElements: true },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    expect(xml).not.toContain('counter(decision)');
    expect(xml).toContain('<a:t>1. </a:t>');
    expect(xml).toContain('<a:t>2. </a:t>');
  }, 40000);

  it('resolves nested counters with separators and counter styles', async () => {
    const html = `
      <!doctype html>
      <html>
      <head>
      <style>
        body { margin: 0; }
        .slide { width: 1920px; height: 1080px; position: relative; background: #fff; }
        .outline { counter-reset: item; font: 400 30px/1.3 Arial; }
        .branch { counter-increment: item; }
        .branch::before { content: counter(item, upper-roman) ". "; font-weight: 700; }
        .children { counter-reset: item; margin-left: 40px; }
        .leaf { counter-increment: item; }
        .leaf::before { content: counters(item, ".", upper-roman) " "; font-weight: 700; }
      </style>
      </head>
      <body>
        <div class="slide">
          <div class="outline">
            <div class="branch">Rahmen
              <div class="children"><div class="leaf">Wirkung</div></div>
            </div>
            <div class="branch">Entscheidung</div>
          </div>
        </div>
      </body>
      </html>
    `;

    const buffer = await exportHtmlToPptx(html, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, includePseudoElements: true },
    });
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    expect(xml).not.toContain('counter(');
    expect(xml).not.toContain('counters(');
    expect(xml).toContain('<a:t>I. </a:t>');
    expect(xml).toContain('<a:t>I.I </a:t>');
    expect(xml).toContain('<a:t>II. </a:t>');
  }, 40000);
});
