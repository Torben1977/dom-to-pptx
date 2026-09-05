# Supported CSS & HTML

This document lists the common CSS features, Tailwind-like utility classes, and HTML elements that dom-to-pptx understands and maps to PowerPoint shapes/text.

The browser is the layout authority, but PowerPoint is not a browser. A computed
browser rectangle alone is not sufficient for every CSS construct: flowing text,
clipping, transforms, and paint order may require several PowerPoint objects or a
fidelity fallback. The converter therefore has three explicit support levels:

1. **Native and editable:** text, common boxes, basic borders, simple images/SVG,
   and fixed flex/grid layouts are mapped to editable PowerPoint objects.
2. **Fidelity fallback:** a visual subtree may be emitted as vector or raster media
   when PowerPoint has no equivalent editable primitive.
3. **Unsupported boundary:** constructs that cannot be represented reliably are
   documented and covered by executable expected-failure tests. They must not be
   assumed to work merely because the browser renders them.

`exportToPptx(..., { boundaryPolicy })` makes that distinction executable:

- `ignore` keeps the best-effort editable mapper and is the default;
- `error` rejects unsupported conversion boundaries with a structured
  `DOM_TO_PPTX_UNSUPPORTED_BOUNDARY` finding. Use this for controlled presentation
  HTML so the authoring layer can repair the source;
- `rasterize` replaces only the smallest affected unsupported subtree with a PNG.
  Use this fidelity fallback for external HTML when visual preservation matters
  more than editability of that subtree.

Simple axis-aligned clipping of an empty solid rectangle remains native and
editable in every mode.

Simple CSS multi-column regions remain editable when each direct block occupies
one browser column. Browser-balanced direct text or a block fragmented across
columns is rejected in `error` mode and isolated as one raster subtree in
`rasterize` mode.

## Supported HTML elements

- div, span, p, h1-h6
- img, svg
- ul, ol, li
- a
- button
- section, article, header, footer
- input (text), textarea (simple text extraction)
- figure, figcaption

## Supported CSS properties (rendered visually)

- background-color, background-image (linear-gradient)
- background-position, background-size (basic handling in gradients)
- color, opacity
- border, border-_-color, border-_-width, border-radius (per-corner)
- box-shadow (outer shadows mapped to PPTX outer shadows)
- filter: blur() (soft-edge rendering via SVG)
- backdrop-filter: blur() (simulated via html2canvas snapshot)
- transform: rotate() (extraction of rotation angle)
- display, position, width, height, padding, margin
- CSS stacking contexts and `z-index` paint order for positioned elements,
  flex/grid items, opacity, and transformed wrappers
- simple left/right sibling floats when all affected text lines share one
  rectangular line region
- text-align, vertical-align, text-transform
- white-space (`normal`/`nowrap` collapse whitespace; `pre`/`pre-wrap`/`pre-line` preserve author line breaks, and `pre`/`pre-wrap` also preserve indentation/spaces)
- font-family, font-size, font-weight, font-style, line-height
- animations & transitions (20+ entrance/exit animations like `fade-in`, `zoom-in`, `fly-in`, `wipe-in`; 70+ slide transitions; custom delays, durations, trigger sequencing, and character/paragraph reveals)

## Common utility/Tailwind-like classes (recognized by visual result)

These classes are examples; dom-to-pptx reads computed styles, so any combination that results in the same computed value will be supported.

- `rounded`, `rounded-sm`, `rounded-md`, `rounded-lg`, `rounded-xl`, `rounded-full`, `rounded-tr-*`, `rounded-bl-full`, etc.
- `bg-white`, `bg-slate-50`, `bg-indigo-50`, `bg-gradient-to-r`, `from-indigo-400`, `to-cyan-400`, etc. (linear-gradients are parsed)
- `shadow`, `shadow-md`, `shadow-lg`, `shadow-2xl` (box-shadow)
- `flex`, `grid`, `items-center`, `justify-center`, `gap-*`
- `p-4`, `px-6`, `py-2`, `m-4`
- `w-*`, `h-*` (fixed pixel/percentage/wrappers — computed width/height are used)
- `text-xs`, `text-sm`, `text-lg`, `font-bold`, `uppercase`, `italic`, `tracking-wide`
- `fade-in`, `fly-in`, `wipe-out`, `animate-duration-[500]`, `animate-delay-[200]`, `animate-trigger-after` (animation utility classes)
- `slide-transition-fade`, `slide-transition-push`, `transition-dur-[1000]` (transition utility classes)

## Limitations

- Infinite-loop, interaction-triggered, or non-whitelisted CSS animations/transitions are not exported. Only the whitelisted 20+ slide element animations and 70+ slide transitions are translated into PowerPoint motion effects; all others will fallback to their static computed layout.
- Some advanced CSS features (CSS variables used as colors, filters beyond blur) may not map 1:1.
- For images to be processed via canvas (rounded images), the source must be CORS-accessible (`Access-Control-Allow-Origin` header) or the image will be skipped or rendered as-is.

### Recommended Patterns & Best Practices

To achieve the highest fidelity and most reliable rendering in PowerPoint, consider adopting these HTML/CSS patterns:

- **Layouts (Tables vs. Grid/Flex):** While native HTML `<table>` elements are supported (mapping to PptxGenJS native tables), native tables in PowerPoint can be structurally rigid. For absolute layout control, perfect border-radius, and guaranteed visual consistency, **prefer utilizing a `div` structure with `display: grid` or `display: flex`**. These containers dynamically transform into crisp, independent PowerPoint shapes.
- **Images and Backgrounds:**
  - `<img src="...">` tags support the common `object-fit` and
    `object-position` cases. Complex crop/radius combinations require the media
    fidelity tests and are not described as pixel-perfect by default.
  - Static PNG, JPEG and WebP sources are embedded without re-encoding when
    they fill their picture bounds without cropping, padding or rounded masks
    and fit within the existing 2x target-resolution ceiling. Oversized originals
    still downsample, so small image tiles do not embed unnecessarily large files.
    CORS-readable URLs and blob URLs use their response MIME type. Sources that
    require rasterization retain the existing 2x output ceiling but do not exceed
    their available pixel density after fitting. SVG keeps the vector-source
    rasterization resolution instead of being capped to intrinsic dimensions.
  - Rasterized WebP remains WebP at browser encoder quality 1, preserving alpha.
    This is maximum-quality **lossy** re-encoding, not lossless WebP. PNG inputs
    and other raster fallbacks remain PNG. A browser without a WebP encoder may
    return PNG, whose actual MIME header is passed to the PPTX writer.
  - Direct WebP was checked in PowerPoint for Mac 16.112.3 and LibreOffice
    26.2.5.2. Windows, older Office versions, and animated WebP are not certified
    by these checks. Run `npm run test:office-roundtrip` on the target renderer
    to test the exported crop, alpha and opacity in addition to package structure.
  - CSS `background-image: url(...)` is also natively parsed. It correctly handles `background-size` (cover/contain) and translates them into matching image crop parameters in PPTX.
  - CSS `background-image: linear-gradient(...)` transforms into pure vector SVG gradients without requiring rasterization.
- **Writing Modes:** Modern CSS `writing-mode` (`vertical-rl`, `vertical-lr`) properties are supported. Combine them with `text-orientation: upright` to natively tap into PowerPoint's Stacked Vertical Text engine, or leave defaults to map to standard East-Asian rotated text layouts.

If a style or element is critical and you find it not behaving as expected, open an issue with a minimal repro and I'll add support or provide a workaround.

## Known renderer boundaries

The following browser constructs are deliberately tracked as renderer boundaries
instead of being hidden behind slide-specific repairs:

- text that first flows beside a float and then continues below it (a
  non-rectangular line region);
- `overflow: hidden` / `clip` for editable text, media, or transformed children
  under the default best-effort mapper. Select `boundaryPolicy: error` or
  `boundaryPolicy: rasterize` to make this boundary explicit;
- cumulative ancestor transforms, non-default `transform-origin`, scale, skew,
  and perspective;
- mixed Shadow-DOM content containing both editable text and visual elements;
- nested lists, advanced list markers, and arbitrary CSS counters;
- browser-balanced or fragmented CSS multi-column layout, `shape-outside`,
  masks, and complex `clip-path`;
- blend modes, backdrop effects beyond the documented fallback, video, iframe,
  WebGL, interaction, and form behavior.

For presentation HTML, prefer fixed slide dimensions, explicit fonts and sizes,
simple DOM ownership (one editable object per element), and layout structures that
have a direct PowerPoint representation. Unsupported interactive constructs should
be rejected or flattened intentionally, not approximated silently.
