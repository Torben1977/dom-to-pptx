// Runs two builds of the converter over a folder of benchmark decks and reports,
// slide by slide, what changed and whether each change brings Office closer to
// the browser. For judging a converter fix against real decks before release.
//
//   node scripts/compare-benchmark-decks.mjs --base <checkout> --fix <checkout> \
//     --runs <org-graph>/tests/agents/evals/runs --out <dir> [--filter <text>] [--concurrency 3]
//
// Both checkouts need `npm run build` (the exporter injects dist/dom-to-pptx.bundle.js).
// Each run folder must hold presentation/deck.html, the adapter output the
// converter gets in production. Needs soffice and pdftotext on the PATH.
//
// --renderer powerpoint renders with Microsoft PowerPoint instead of LibreOffice
// (macOS; the terminal app needs the Automation permission for PowerPoint). It
// opens one deck at a time in PowerPoint, exports it as PDF and closes it
// unsaved; open presentations are left alone.
//
// Per slide it reports four things:
//   - XML changes between the builds: text frame geometry and insets, table cell
//     fills and margins, row heights;
//   - the fidelity findings of both builds (the defect kinds of the Treue-Test);
//   - per word, how far Office puts it from the browser, summed over the words
//     whose position changed, apart for blocks of one line and of several: the
//     converter places them by different rules, and a fix to one can hide in the
//     sum of both;
//   - text frames an opaque shape or picture painted after them covers. pdftotext
//     finds covered text all the same, so no word measurement notices it. A
//     picture covers where its own pixels are opaque, sampled in a browser.
//
// The summary adds, over all slides, each build's mean offset per word.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { execFileSync } from 'node:child_process';
import puppeteer from 'puppeteer';
import JSZip from 'jszip';
import {
  analyzePage,
  assignWords,
  convertToPdf,
  groupLines,
  measureBrowserPages,
  parseOfficeWords,
} from '../src/__tests__/helpers/office-fidelity.js';

const args = Object.fromEntries(
  process.argv
    .slice(2)
    .reduce(
      (pairs, value, index, all) => (value.startsWith('--') ? [...pairs, [value.slice(2), all[index + 1]]] : pairs),
      []
    )
);
for (const required of ['base', 'fix', 'runs', 'out']) {
  if (!args[required]) {
    console.error(`missing --${required}; see the header of ${path.basename(process.argv[1])}`);
    process.exit(2);
  }
}
const concurrency = Number(args.concurrency || 3);
if (!['libreoffice', 'powerpoint'].includes(args.renderer || 'libreoffice')) {
  console.error(`--renderer must be libreoffice or powerpoint, not ${args.renderer}`);
  process.exit(2);
}

// PowerPoint is one application, so its conversions queue up; PowerPoint is
// sandboxed and reads and writes only inside its own container.
const POWERPOINT_DIR = path.join(
  os.homedir(),
  'Library/Containers/com.microsoft.Powerpoint/Data/Documents/dom-to-pptx-compare'
);
let powerPointQueue = Promise.resolve();
// A PowerPoint that stops answering (a dialog waiting for the user, say) stays
// stuck; everything after it would pile up behind it. The first timeout ends
// all further conversions instead.
let powerPointStuck = false;
function convertWithPowerPoint(pptxPath, outDir) {
  const convert = () => {
    if (powerPointStuck) throw new Error('PowerPoint stopped answering earlier in this run; not converting');
    fs.mkdirSync(POWERPOINT_DIR, { recursive: true });
    const name = `${path.basename(path.dirname(outDir))}-${path.basename(outDir)}.pptx`.replace(/[^\w.-]/g, '_');
    const source = path.join(POWERPOINT_DIR, name);
    const pdf = source.replace(/\.pptx$/, '.pdf');
    fs.copyFileSync(pptxPath, source);
    try {
      execFileSync(
        'osascript',
        [
          '-e',
          [
            'with timeout of 90 seconds',
            'tell application "Microsoft PowerPoint"',
            `  open POSIX file "${source}"`,
            '  repeat 100 times',
            `    if exists presentation "${name}" then exit repeat`,
            '    delay 0.2',
            '  end repeat',
            `  set p to presentation "${name}"`,
            `  save p in POSIX file "${pdf}" as save as PDF`,
            '  close p saving no',
            'end tell',
            'end timeout',
          ].join('\n'),
        ],
        { timeout: 120_000 }
      );
      fs.copyFileSync(pdf, path.join(outDir, 'deck.pdf'));
    } catch (error) {
      if (error.code === 'ETIMEDOUT' || /-1712/.test(String(error.stderr || error.message))) powerPointStuck = true;
      throw error;
    } finally {
      fs.rmSync(source, { force: true });
      fs.rmSync(pdf, { force: true });
    }
  };
  const done = powerPointQueue.then(convert);
  powerPointQueue = done.catch(() => {});
  return done;
}
const render = (pptxPath, outDir) =>
  args.renderer === 'powerpoint' ? convertWithPowerPoint(pptxPath, outDir) : convertToPdf(pptxPath, outDir);
const CANVAS = { width: 1280, height: 720 };
const VARIANTS = ['base', 'fix'];
const exporters = Object.fromEntries(
  await Promise.all(
    VARIANTS.map(async (variant) => [
      variant,
      (await import(path.resolve(args[variant], 'src/node-exporter.js'))).exportHtmlToPptx,
    ])
  )
);
const executablePath = await puppeteer.executablePath();

// Whether a picture hides what lies beneath depends on its pixels at the spot,
// not on its box: an icon or a rounded border drawn as a picture is mostly
// transparent. One browser page samples them for the whole run.
const sampler = await puppeteer.launch({ executablePath, headless: true, args: ['--no-sandbox'] });
const samplerPage = await sampler.newPage();

/** Alpha (0–1) of a picture at (u, v) of its box, drawn as PowerPoint stretches it into the box. */
function pictureAlphaAt({ dataUrl, u, v, boxWidth, boxHeight, crop }) {
  return samplerPage.evaluate(
    async (source) => {
      const image = new Image();
      image.src = source.dataUrl;
      await image.decode();
      const scale = Math.min(1, 512 / Math.max(source.boxWidth, source.boxHeight));
      const [width, height] = [source.boxWidth, source.boxHeight].map((size) => Math.max(1, Math.round(size * scale)));
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext('2d');
      const [naturalWidth, naturalHeight] = [image.naturalWidth || width, image.naturalHeight || height];
      const { l, t, r, b } = source.crop;
      context.drawImage(
        image,
        l * naturalWidth,
        t * naturalHeight,
        (1 - l - r) * naturalWidth,
        (1 - t - b) * naturalHeight,
        0,
        0,
        width,
        height
      );
      const x = Math.min(width - 1, Math.floor(source.u * width));
      const y = Math.min(height - 1, Math.floor(source.v * height));
      return context.getImageData(x, y, 1, 1).data[3] / 255;
    },
    { dataUrl, u, v, boxWidth, boxHeight, crop }
  );
}

const MEDIA_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml' };
const OPAQUE_ALPHA = 0.95;
const EMU_PER_PX = 9525;

const runs = fs
  .readdirSync(args.runs)
  .filter((run) => fs.existsSync(path.join(args.runs, run, 'presentation/deck.html')))
  .filter((run) => !args.filter || run.includes(args.filter))
  .sort();

function xmlFacts(xml) {
  const shapes = Array.from(xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g), ([, sp]) => {
    const off = sp.match(/<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/);
    const inset = sp.match(/<a:bodyPr[^>]*?lIns="(\d+)"[^>]*?rIns="(\d+)"/);
    const text = Array.from(sp.matchAll(/<a:t>([^<]*)<\/a:t>/g), (run) => run[1]).join('');
    return { text: text.slice(0, 50), off: off?.slice(1).join(','), inset: inset?.slice(1).join(',') };
  }).filter((shape) => shape.text);
  const cells = [];
  for (const [, table] of xml.matchAll(/<a:tbl>([\s\S]*?)<\/a:tbl>/g)) {
    for (const [, rowHeight, row] of table.matchAll(/<a:tr h="(\d+)">([\s\S]*?)<\/a:tr>/g)) {
      for (const [, cell] of row.matchAll(/<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/g)) {
        const props = cell.match(/<a:tcPr([^>]*)>([\s\S]*?)<\/a:tcPr>/);
        const own = props ? props[2].replace(/<a:ln[LRTB][\s\S]*?<\/a:ln[LRTB]>/g, '') : '';
        const fill = own.match(/<a:solidFill><a:srgbClr val="(\w+)"(?:><a:alpha val="(\d+)"\/>)?/);
        const margin = (side) => props?.[1].match(new RegExp(`${side}="(\\d+)"`))?.[1] ?? '0';
        cells.push({
          text: Array.from(cell.matchAll(/<a:t>([^<]*)<\/a:t>/g), (run) => run[1])
            .join('')
            .slice(0, 40),
          rowHeight,
          margin: `${margin('marL')},${margin('marR')}`,
          fill: fill ? `${fill[1]}${fill[2] === undefined ? '' : `@${fill[2]}`}` : 'none',
        });
      }
    }
  }
  return { shapes, cells };
}

// Text frames that an opaque shape or picture painted after them covers at their
// centre.
async function coveredTexts(xml, zip, slideNumber) {
  const rels = (await zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)?.async('string')) || '';
  const targets = new Map(
    Array.from(rels.matchAll(/<Relationship\b([^>]*)\/>/g), ([, attributes]) => [
      attributes.match(/\bId="([^"]+)"/)?.[1],
      path.posix.normalize(path.posix.join('ppt/slides', attributes.match(/\bTarget="([^"]+)"/)?.[1] || '')),
    ])
  );
  const shapes = Array.from(xml.matchAll(/<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g), ([, kind, body]) => {
    const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!off) return null;
    const [x, y, w, h] = off.slice(1).map(Number);
    const text = Array.from(body.matchAll(/<a:t>([^<]*)<\/a:t>/g), (run) => run[1]).join('');
    return { kind, body, x, y, w, h, text: text.trim() };
  }).filter(Boolean);

  const alphaOf = (colors) =>
    Array.from(colors.matchAll(/<a:srgbClr val="\w+"(?:\/>|>([\s\S]*?)<\/a:srgbClr>)/g), (color) =>
      Number(color[1]?.match(/<a:alpha val="(\d+)"/)?.[1] ?? 100000)
    );
  async function covers(over, px, py) {
    if (over.w <= 0 || over.h <= 0 || px < over.x || px > over.x + over.w || py < over.y || py > over.y + over.h) {
      return false;
    }
    if (over.kind === 'sp') {
      // The fill, not the outline: a solid colour or a gradient whose every stop is opaque.
      const spPr = (over.body.match(/<p:spPr\b[\s\S]*?<\/p:spPr>/)?.[0] || '').replace(/<a:ln\b[\s\S]*?<\/a:ln>/g, '');
      const fill = spPr.match(/<a:(solidFill|gradFill)\b[\s\S]*?<\/a:\1>/)?.[0];
      const alphas = fill ? alphaOf(fill) : [];
      return alphas.length > 0 && alphas.every((alpha) => alpha >= OPAQUE_ALPHA * 100000);
    }
    const embed =
      over.body.match(/<asvg:svgBlip\b[^>]*r:embed="([^"]+)"/)?.[1] ||
      over.body.match(/<a:blip\b[^>]*r:embed="([^"]+)"/)?.[1];
    const media = embed && zip.file(targets.get(embed) || '');
    // A picture whose data cannot be read counts as opaque, as a box would.
    if (!media) return true;
    const type = MEDIA_TYPES[path.extname(media.name).slice(1).toLowerCase()] || 'image/png';
    const srcRect = over.body.match(/<a:srcRect\b([^>]*)\/>/)?.[1] || '';
    const edge = (side) => Number(srcRect.match(new RegExp(`\\b${side}="(-?\\d+)"`))?.[1] || 0) / 100000;
    const amount = Number(over.body.match(/<a:alphaModFix amt="(\d+)"/)?.[1] ?? 100000) / 100000;
    const alpha = await pictureAlphaAt({
      dataUrl: `data:${type};base64,${await media.async('base64')}`,
      u: (px - over.x) / over.w,
      v: (py - over.y) / over.h,
      boxWidth: over.w / EMU_PER_PX,
      boxHeight: over.h / EMU_PER_PX,
      crop: { l: edge('l'), t: edge('t'), r: edge('r'), b: edge('b') },
    });
    return alpha * amount >= OPAQUE_ALPHA;
  }

  const covered = [];
  for (const [index, shape] of shapes.entries()) {
    if (!shape.text) continue;
    const [cx, cy] = [shape.x + shape.w / 2, shape.y + shape.h / 2];
    for (const over of shapes.slice(index + 1)) {
      if (await covers(over, cx, cy)) {
        covered.push(shape.text.slice(0, 50));
        break;
      }
    }
  }
  return covered;
}

// Frames pair up by their text, and frames that share a text with the nearest
// one, so a frame that only moves in the paint order keeps its partner.
function pairFrames(before, after) {
  const position = (frame) => (frame.off || '0,0').split(',').map(Number);
  const candidates = [];
  before.forEach((frame, i) =>
    after.forEach((other, j) => {
      if (frame.text !== other.text) return;
      const [[x1, y1], [x2, y2]] = [position(frame), position(other)];
      candidates.push({ i, j, distance: Math.hypot(x1 - x2, y1 - y2) });
    })
  );
  const [pairedBefore, pairedAfter, pairs] = [new Set(), new Set(), []];
  for (const { i, j } of candidates.sort((a, b) => a.distance - b.distance)) {
    if (pairedBefore.has(i) || pairedAfter.has(j)) continue;
    pairedBefore.add(i);
    pairedAfter.add(j);
    pairs.push([before[i], after[j]]);
  }
  return {
    pairs,
    removed: before.filter((_, i) => !pairedBefore.has(i)),
    added: after.filter((_, j) => !pairedAfter.has(j)),
  };
}

function xmlChanges(base, fix) {
  const changes = [];
  if (base.cells.length !== fix.cells.length)
    changes.push(`structure: ${base.cells.length} -> ${fix.cells.length} table cells`);
  const frames = pairFrames(base.shapes, fix.shapes);
  for (const [before, after] of frames.pairs) {
    for (const key of ['off', 'inset']) {
      if (before[key] !== after[key]) changes.push(`frame-${key}: "${before.text}" ${before[key]} -> ${after[key]}`);
    }
  }
  for (const frame of frames.removed) changes.push(`frame-removed: "${frame.text}"`);
  for (const frame of frames.added) changes.push(`frame-added: "${frame.text}"`);
  if (base.shapes.map((shape) => shape.text).join('\n') !== fix.shapes.map((shape) => shape.text).join('\n'))
    changes.push('order: text frames painted in another order');
  base.cells.forEach((cell, index) => {
    for (const key of ['fill', 'margin', 'rowHeight']) {
      if (fix.cells[index] && cell[key] !== fix.cells[index][key]) {
        changes.push(`cell-${key}: "${cell.text}" ${cell[key]} -> ${fix.cells[index][key]}`);
      }
    }
  });
  return changes;
}

// A finding without its numbers: the same drift, a few points smaller, is the
// same finding; how far things moved is what the word deviations measure.
const findingKey = (finding) => finding.replace(/-?\d+(\.\d+)?/g, '#');

// Per word, Office position minus browser position; null where Office has no
// such word. Taken as it is: a deviation relative to the slide's median misled
// wherever a fix moved only part of a slide's frames, because the median moved
// with them (a clear improvement read as 240 slides worse).
function wordDeviations(browserWords, officeWords) {
  const pairs = assignWords(browserWords, officeWords);
  return browserWords.map((word, index) =>
    pairs.has(index)
      ? { dy: officeWords[pairs.get(index)].top - word.top, dx: officeWords[pairs.get(index)].x - word.x }
      : null
  );
}

const LINE_KINDS = ['single', 'multi'];

// Per word, whether its block holds one line or several in the browser.
function lineKinds(words) {
  const blocks = new Map();
  for (const word of words) blocks.set(word.block, [...(blocks.get(word.block) || []), word]);
  const kinds = new Map(
    [...blocks].map(([block, blockWords]) => [block, groupLines(blockWords).length > 1 ? 'multi' : 'single'])
  );
  return words.map((word) => kinds.get(word.block));
}

const offset = (deviation) => Math.abs(deviation.dy) + Math.abs(deviation.dx);

async function compareRun(run) {
  const outDir = path.join(args.out, run);
  fs.mkdirSync(outDir, { recursive: true });
  const deck = fs.readFileSync(path.join(args.runs, run, 'presentation/deck.html'), 'utf8');
  let slide = 0;
  // Every slide is one probe, so every word of it is measured.
  const probedPath = path.join(outDir, 'deck.probed.html');
  fs.writeFileSync(
    probedPath,
    deck.replace(/<section\b(?=[^>]*\bclass="slide\b)/g, () => `<section data-probe="slide-${++slide}"`)
  );

  const { pages } = await measureBrowserPages(executablePath, probedPath, CANVAS);
  const variants = {};
  for (const variant of VARIANTS) {
    const variantDir = path.join(outDir, variant);
    fs.mkdirSync(variantDir, { recursive: true });
    const buffer = await exporters[variant](probedPath, {
      selector: '.slide',
      pptxOptions: { width: 13.333333, height: 7.5, autoEmbedFonts: false, boundaryPolicy: 'rasterize' },
    });
    const pptxPath = path.join(variantDir, 'deck.pptx');
    fs.writeFileSync(pptxPath, buffer);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await Promise.all(
      pages.map((_, index) => zip.file(`ppt/slides/slide${index + 1}.xml`).async('string'))
    );
    await render(pptxPath, variantDir);
    const bboxPath = path.join(variantDir, 'deck.bbox.html');
    execFileSync('pdftotext', ['-bbox', path.join(variantDir, 'deck.pdf'), bboxPath], { stdio: 'pipe' });
    const office = parseOfficeWords(fs.readFileSync(bboxPath, 'utf8'));
    variants[variant] = pages.map((page, index) => ({
      xml: xmlFacts(xml[index]),
      findings: analyzePage(page, office[index] || [], xml[index], {}).map(({ kind, detail }) => `${kind}: ${detail}`),
      deviations: wordDeviations(page.words, office[index] || []),
    }));
    for (const [index, slideFacts] of variants[variant].entries()) {
      slideFacts.covered = await coveredTexts(xml[index], zip, index + 1);
    }
  }

  const slides = pages.map((page, index) => {
    const [base, fix] = VARIANTS.map((variant) => variants[variant][index]);
    const kinds = lineKinds(page.words);
    const words = page.words.map((word, wordIndex) => ({
      word: word.text,
      kind: kinds[wordIndex],
      base: base.deviations[wordIndex],
      fix: fix.deviations[wordIndex],
    }));
    const moved = words.filter(({ base: before, fix: after }) =>
      before && after ? Math.abs(before.dy - after.dy) > 0.3 || Math.abs(before.dx - after.dx) > 0.3 : before !== after
    );
    const [baseKeys, fixKeys] = [base, fix].map((variant) => new Set(variant.findings.map(findingKey)));
    const sum = (variant, kind) =>
      Number(
        moved
          .filter((word) => word.kind === kind && word[variant])
          .reduce((total, word) => total + offset(word[variant]), 0)
          .toFixed(1)
      );
    return {
      slide: index + 1,
      xmlChanges: xmlChanges(base.xml, fix.xml),
      findingsGone: base.findings.filter((finding) => !fixKeys.has(findingKey(finding))),
      findingsAdded: fix.findings.filter((finding) => !baseKeys.has(findingKey(finding))),
      movedWords: moved.length,
      wordsLost: moved.filter((word) => word.base && !word.fix).map((word) => word.word),
      wordsFound: moved.filter((word) => !word.base && word.fix).map((word) => word.word),
      deviationPt: Object.fromEntries(
        LINE_KINDS.map((kind) => [kind, Object.fromEntries(VARIANTS.map((variant) => [variant, sum(variant, kind)]))])
      ),
      covered: { base: base.covered, fix: fix.covered },
      textCovered: fix.covered.filter((text) => !base.covered.includes(text)),
      textUncovered: base.covered.filter((text) => !fix.covered.includes(text)),
      // The words both builds place, for the summary's mean offset per word: a
      // word only one build finds would skew one mean and not the other.
      offsets: Object.fromEntries(
        LINE_KINDS.map((kind) => {
          const shared = words.filter((word) => word.kind === kind && word.base && word.fix);
          return [
            kind,
            Object.fromEntries(VARIANTS.map((variant) => [variant, shared.map((word) => offset(word[variant]))])),
          ];
        })
      ),
    };
  });
  const result = { run, slides };
  fs.writeFileSync(path.join(outDir, 'result.json'), JSON.stringify(result, null, 2));
  return result;
}

function verdict(slide) {
  const deviations = LINE_KINDS.map((kind) => slide.deviationPt[kind]);
  const worse =
    slide.findingsAdded.length > 0 ||
    slide.wordsLost.length > 0 ||
    slide.textCovered.length > 0 ||
    deviations.some(({ base, fix }) => fix > base + 0.5);
  const better =
    slide.findingsGone.length > 0 ||
    slide.wordsFound.length > 0 ||
    slide.textUncovered.length > 0 ||
    deviations.some(({ base, fix }) => fix < base - 0.5);
  return worse && better ? 'mixed' : worse ? 'worse' : better ? 'better' : 'same';
}

const queue = [...runs];
const results = [];
const failures = [];
await Promise.all(
  Array.from({ length: concurrency }, async () => {
    while (queue.length) {
      const run = queue.shift();
      try {
        const result = await compareRun(run);
        results.push(result);
        const changed = result.slides.filter((slide) => verdict(slide) !== 'same' || slide.xmlChanges.length);
        console.log(`${run}: ${result.slides.length} slides, ${changed.length} changed`);
      } catch (error) {
        failures.push({ run, error: String(error?.stack || error) });
        console.log(`${run}: FAILED ${error?.message}`);
      }
    }
  })
);

const allSlides = results.flatMap((result) =>
  result.slides.map((slide) => ({ run: result.run, ...slide, verdict: verdict(slide) }))
);
// The word offsets feed the mean below; the list of changed slides leaves them out.
const changed = allSlides
  .filter((slide) => slide.verdict !== 'same' || slide.xmlChanges.length)
  .map(({ offsets, ...slide }) => slide);
const mean = (values) =>
  values.length ? Number((values.reduce((a, b) => a + b, 0) / values.length).toFixed(2)) : null;
const meanOffsetPt = Object.fromEntries(
  LINE_KINDS.map((kind) => [
    kind,
    Object.fromEntries(
      VARIANTS.map((variant) => [variant, mean(allSlides.flatMap((slide) => slide.offsets[kind][variant]))])
    ),
  ])
);
const summary = {
  runs: results.length,
  slides: allSlides.length,
  failures,
  meanOffsetPt,
  verdicts: changed.reduce((counts, slide) => ({ ...counts, [slide.verdict]: (counts[slide.verdict] || 0) + 1 }), {}),
  changed,
};
fs.writeFileSync(path.join(args.out, 'summary.json'), JSON.stringify(summary, null, 2));
console.log(`\n${summary.runs} runs, ${summary.slides} slides, ${changed.length} changed, ${failures.length} failed`);
for (const kind of LINE_KINDS) {
  const { base, fix } = meanOffsetPt[kind];
  console.log(`mean offset per word placed by both builds, ${kind}-line blocks: ${base} -> ${fix} pt`);
}
const pair = ({ base, fix }) => `${base} -> ${fix}`;
for (const slide of changed) {
  console.log(
    `${slide.verdict.padEnd(6)} ${slide.run} s${slide.slide}: deviation single ${pair(slide.deviationPt.single)} pt, ` +
      `multi ${pair(slide.deviationPt.multi)} pt, findings -${slide.findingsGone.length}/+${slide.findingsAdded.length}, ` +
      `covered text -${slide.textUncovered.length}/+${slide.textCovered.length}, ${slide.xmlChanges.length} XML changes`
  );
}
await sampler.close();
