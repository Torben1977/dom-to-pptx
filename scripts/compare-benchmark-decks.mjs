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
//     finds covered text all the same, so no word measurement notices it.
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

const runs = fs
  .readdirSync(args.runs)
  .filter((run) => fs.existsSync(path.join(args.runs, run, 'presentation/deck.html')))
  .filter((run) => !args.filter || run.includes(args.filter))
  .sort();

function xmlFacts(xml) {
  const shapes = Array.from(xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g), ([, sp]) => {
    const off = sp.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/);
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
function coveredTexts(xml) {
  const shapes = Array.from(xml.matchAll(/<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g), ([, kind, body]) => {
    const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/);
    if (!off) return null;
    const [x, y, w, h] = off.slice(1).map(Number);
    const fill = body
      .match(/<p:spPr\b[\s\S]*?<\/p:spPr>/)?.[0]
      .match(/<a:solidFill><a:srgbClr val="\w+"(?:\/>|>([\s\S]*?)<\/a:srgbClr>)/);
    const alpha = Number(fill?.[1]?.match(/<a:alpha val="(\d+)"/)?.[1] ?? 100000);
    const text = Array.from(body.matchAll(/<a:t>([^<]*)<\/a:t>/g), (run) => run[1]).join('');
    return { x, y, w, h, text: text.trim(), opaque: kind === 'pic' || (Boolean(fill) && alpha >= 95000) };
  }).filter(Boolean);
  return shapes
    .filter((shape, index) => {
      if (!shape.text) return false;
      const [cx, cy] = [shape.x + shape.w / 2, shape.y + shape.h / 2];
      return shapes
        .slice(index + 1)
        .some((over) => over.opaque && cx >= over.x && cx <= over.x + over.w && cy >= over.y && cy <= over.y + over.h);
    })
    .map((shape) => shape.text.slice(0, 50));
}

function xmlChanges(base, fix) {
  const changes = [];
  if (base.shapes.length !== fix.shapes.length)
    changes.push(`structure: ${base.shapes.length} -> ${fix.shapes.length} text frames`);
  if (base.cells.length !== fix.cells.length)
    changes.push(`structure: ${base.cells.length} -> ${fix.cells.length} table cells`);
  const compare = (kind, before, after, keys) =>
    before.forEach((item, index) => {
      for (const key of keys) {
        if (after[index] && item[key] !== after[index][key]) {
          changes.push(`${kind}-${key}: "${item.text}" ${item[key]} -> ${after[index][key]}`);
        }
      }
    });
  // Frames pair up by their text, so one that only moves in the paint order
  // does not shift every frame after it onto a stranger.
  const byText = (shapes) => [...shapes].sort((a, b) => a.text.localeCompare(b.text));
  compare('frame', byText(base.shapes), byText(fix.shapes), ['off', 'inset']);
  if (base.shapes.map((shape) => shape.text).join('\n') !== fix.shapes.map((shape) => shape.text).join('\n'))
    changes.push('order: text frames painted in another order');
  compare('cell', base.cells, fix.cells, ['fill', 'margin', 'rowHeight']);
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
      covered: coveredTexts(xml[index]),
    }));
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
      textCovered: fix.covered.filter((text) => !base.covered.includes(text)),
      textUncovered: base.covered.filter((text) => !fix.covered.includes(text)),
      // Every matched word, for the summary's mean offset per word.
      offsets: Object.fromEntries(
        LINE_KINDS.map((kind) => [
          kind,
          Object.fromEntries(
            VARIANTS.map((variant) => [
              variant,
              words.filter((word) => word.kind === kind && word[variant]).map((word) => offset(word[variant])),
            ])
          ),
        ])
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
  console.log(`mean offset per word, ${kind}-line blocks: ${base} -> ${fix} pt`);
}
const pair = ({ base, fix }) => `${base} -> ${fix}`;
for (const slide of changed) {
  console.log(
    `${slide.verdict.padEnd(6)} ${slide.run} s${slide.slide}: deviation single ${pair(slide.deviationPt.single)} pt, ` +
      `multi ${pair(slide.deviationPt.multi)} pt, findings -${slide.findingsGone.length}/+${slide.findingsAdded.length}, ` +
      `covered text -${slide.textUncovered.length}/+${slide.textCovered.length}, ${slide.xmlChanges.length} XML changes`
  );
}
