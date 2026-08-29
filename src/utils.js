// src/utils.js

// canvas context for color normalization
let _ctx;
function getCtx() {
  if (!_ctx) _ctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true });
  return _ctx;
}

function getTableBorder(style, side, scale, node) {
  const widthStr = style[`border${side}Width`];
  const styleStr = style[`border${side}Style`];
  const colorStr = style[`border${side}Color`];

  const width = parseFloat(widthStr) || 0;
  if (width === 0 || styleStr === 'none' || styleStr === 'hidden') {
    return null;
  }

  let color = parseColor(colorStr, style);
  if (!color.hex || color.opacity === 0) return null;
  color = flattenColor(color, node, false);

  let dash = 'solid';
  if (styleStr === 'dashed') dash = 'dash';
  if (styleStr === 'dotted') dash = 'dot';

  return {
    pt: width * 0.75 * scale, // Convert px to pt
    color: color.hex,
    type: dash,
  };
}

/**
 * Extracts native table data for PptxGenJS.
 */
export function extractTableData(node, scale, pseudoContentByNode = null) {
  const rows = [];
  const colWidths = [];

  const widthInches = (element) => {
    const rectWidth = element.getBoundingClientRect().width;
    const cssWidth = parseFloat(window.getComputedStyle(element).width);
    const px = rectWidth > 0 ? rectWidth : cssWidth;
    return Number.isFinite(px) && px > 0 ? px * (1 / 96) * scale : 0;
  };

  // 1. Derive the rendered column grid. A merged header is not evidence that
  // the underlying columns are equally wide, so prefer explicit <col> tracks
  // and then a fully expanded measured row. Equal splitting is only a fallback.
  const colElements = Array.from(node.children)
    .filter((child) => (child?.tagName || '').toLowerCase() === 'colgroup')
    .flatMap((group) => Array.from(group.children).filter((child) => (child?.tagName || '').toLowerCase() === 'col'));

  const explicitColWidths = [];
  for (const col of colElements) {
    const span = parseInt(col.getAttribute('span')) || 1;
    const totalWidth = widthInches(col);
    const widthPerColumn = totalWidth / span;
    if (widthPerColumn <= 0) {
      explicitColWidths.length = 0;
      break;
    }
    for (let i = 0; i < span; i++) explicitColWidths.push(widthPerColumn);
  }

  const isInsideDisplayNoneGroup = (row) => {
    for (let current = row; current && current !== node; current = current.parentElement) {
      if (window.getComputedStyle(current).display === 'none') return true;
    }
    return false;
  };
  const trList = Array.from(node.querySelectorAll('tr')).filter((row) => {
    const rowStyle = window.getComputedStyle(row);
    return !isInsideDisplayNoneGroup(row) && rowStyle.visibility !== 'collapse';
  });
  if (trList.length > 0) {
    const equations = [];
    let activeRowspans = [];
    let logicalColumnCount = explicitColWidths.length;

    for (const row of trList) {
      const cells = Array.from(row.children).filter((cell) =>
        ['td', 'th'].includes((cell?.tagName || '').toLowerCase())
      );
      const occupied = activeRowspans.map((remaining) => remaining > 0);
      const nextRowspans = activeRowspans.map((remaining) => Math.max(0, remaining - 1));
      let cursor = 0;

      for (const cell of cells) {
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

        while (occupied[cursor]) cursor++;
        while (Array.from({ length: colspan }, (_, offset) => occupied[cursor + offset]).some(Boolean)) {
          cursor++;
          while (occupied[cursor]) cursor++;
        }

        const measuredWidth = widthInches(cell);
        if (measuredWidth > 0) {
          equations.push({ start: cursor, span: colspan, width: measuredWidth });
        }

        for (let offset = 0; offset < colspan; offset++) {
          occupied[cursor + offset] = true;
          if (rowspan > 1) {
            nextRowspans[cursor + offset] = Math.max(nextRowspans[cursor + offset] || 0, rowspan - 1);
          }
        }
        cursor += colspan;
      }

      logicalColumnCount = Math.max(logicalColumnCount, occupied.length, cursor);
      activeRowspans = nextRowspans;
    }

    // A partial <colgroup> defines only the tracks it names. Seed those
    // columns, then complete the remaining grid from the cells the browser
    // actually laid out. Treating a partial list as the whole grid causes
    // PptxGenJS to repeat one width across every implicit column.
    const solvedWidths = Array(logicalColumnCount).fill(null);
    explicitColWidths.forEach((width, index) => {
      solvedWidths[index] = width;
    });
    for (const equation of equations.filter(({ span }) => span === 1)) {
      if (solvedWidths[equation.start] == null) {
        solvedWidths[equation.start] = equation.width;
      }
    }

    // A colspan is a linear constraint over adjacent tracks. Resolve every
    // constraint that has a single unknown after the directly measured cells.
    let changed = true;
    while (changed) {
      changed = false;
      for (const equation of equations) {
        const indexes = Array.from({ length: equation.span }, (_, offset) => equation.start + offset);
        const unknown = indexes.filter((index) => solvedWidths[index] == null);
        if (unknown.length !== 1) continue;
        const knownWidth = indexes.reduce((sum, index) => sum + (solvedWidths[index] || 0), 0);
        const remaining = equation.width - knownWidth;
        if (remaining > 0) {
          solvedWidths[unknown[0]] = remaining;
          changed = true;
        }
      }
    }

    // If only merged cells exist, an equal split is the least-assumptive
    // fallback for the still-underdetermined tracks within that exact span.
    for (const equation of equations) {
      const indexes = Array.from({ length: equation.span }, (_, offset) => equation.start + offset);
      const unknown = indexes.filter((index) => solvedWidths[index] == null);
      if (unknown.length === 0) continue;
      const knownWidth = indexes.reduce((sum, index) => sum + (solvedWidths[index] || 0), 0);
      const share = (equation.width - knownWidth) / unknown.length;
      if (share > 0) unknown.forEach((index) => (solvedWidths[index] = share));
    }

    const measuredTableWidth = widthInches(node);
    const knownTotal = solvedWidths.reduce((sum, width) => sum + (width || 0), 0);
    const unresolved = solvedWidths.filter((width) => width == null).length;
    const fallbackWidth =
      unresolved > 0 && measuredTableWidth > knownTotal
        ? (measuredTableWidth - knownTotal) / unresolved
        : logicalColumnCount > 0
          ? measuredTableWidth / logicalColumnCount
          : 0;

    for (const width of solvedWidths) {
      const resolved = width || fallbackWidth;
      if (resolved > 0) colWidths.push(resolved);
    }
  } else {
    colWidths.push(...explicitColWidths);
  }

  const tableStyle = window.getComputedStyle(node);
  const borderSpacing = tableStyle.borderSpacing.split(' ');
  const hSpace = parseFloat(borderSpacing[0]) || 0;
  const vSpace = parseFloat(borderSpacing[1] || borderSpacing[0]) || 0;
  const hSpacePt = hSpace * 0.75 * scale;
  const vSpacePt = vSpace * 0.75 * scale;

  // 2. Iterate Rows
  trList.forEach((tr) => {
    const rowData = [];
    const cellList = Array.from(tr.children).filter((c) => ['td', 'th'].includes((c?.tagName || '').toLowerCase()));

    cellList.forEach((cell) => {
      const style = window.getComputedStyle(cell);
      const cellParts = collectTextParts(cell, style, scale, null, true, 1, pseudoContentByNode);
      // Fallback to plain text if collectTextParts returns empty/invalid
      const cellText = cellParts && cellParts.length > 0 ? cellParts : cell.innerText.replace(/[\n\r\t]+/g, ' ').trim();

      // A. Text Style
      const textStyle = getTextStyle(style, scale);

      // B. Cell Background
      let bg = parseColor(style.backgroundColor, style);
      if ((!bg.hex || bg.opacity === 0) && style.backgroundImage && style.backgroundImage !== 'none') {
        const fallback = getGradientFallbackColor(style.backgroundImage, style);
        if (fallback) bg = parseColor(fallback, style);
      }
      bg = flattenColor(bg, cell);
      const fill = bg.hex && bg.opacity > 0 ? { color: bg.hex } : null;

      // C. Alignment
      let align = 'left';
      if (style.textAlign === 'center') align = 'center';
      if (style.textAlign === 'right' || style.textAlign === 'end') align = 'right';

      let valign = 'top';
      if (style.verticalAlign === 'middle') valign = 'middle';
      if (style.verticalAlign === 'bottom') valign = 'bottom';

      // D. Padding (Margins in PPTX)
      // CSS Padding px -> PPTX Margin pt
      const padding = getPadding(style, scale);
      // getPadding returns [top, right, bottom, left] in inches relative to scale
      // PptxGenJS expects points (pt) for margin: [t, r, b, l]
      // or discrete properties. Let's use discrete for clarity.
      const margin = [
        padding[0] * 72 + vSpacePt / 2, // top
        padding[1] * 72 + hSpacePt / 2, // right
        padding[2] * 72 + vSpacePt / 2, // bottom
        padding[3] * 72 + hSpacePt / 2, // left
      ];

      // E. Borders
      const borderTop = getTableBorder(style, 'Top', scale, cell);
      const borderRight = getTableBorder(style, 'Right', scale, cell);
      const borderBottom = getTableBorder(style, 'Bottom', scale, cell);
      const borderLeft = getTableBorder(style, 'Left', scale, cell);

      // F. Text Direction
      const writingModeVert = getWritingModeVert(style.writingMode, style.textOrientation);
      const textDirection = mapVertToTextDirection(writingModeVert);

      // G. Construct Cell Object
      rowData.push({
        text: cellText,
        options: {
          color: textStyle.color,
          fontFace: textStyle.fontFace,
          fontSize: textStyle.fontSize,
          bold: textStyle.bold,
          italic: textStyle.italic,
          underline: textStyle.underline,

          fill: fill,
          align: align,
          valign: valign,
          margin: margin,

          rowspan: parseInt(cell.getAttribute('rowspan')) || null,
          colspan: parseInt(cell.getAttribute('colspan')) || null,

          border: [borderTop, borderRight, borderBottom, borderLeft],

          ...(textDirection && { textDirection }),
        },
      });
    });

    if (rowData.length > 0) {
      rows.push(rowData);
    }
  });

  return { rows, colWidths };
}

// Checks if any parent element has overflow: hidden which would clip this element
export function isClippedByParent(node) {
  let parent = node.parentElement;
  while (parent && parent !== document.body) {
    const style = window.getComputedStyle(parent);
    const overflow = style.overflow;
    if (overflow === 'hidden' || overflow === 'clip') {
      return true;
    }
    parent = parent.parentElement;
  }
  return false;
}

// Helper to save gradient text
// Helper to save gradient text: extracts the first color from a gradient string
export function getGradientFallbackColor(bgImage, style) {
  if (!bgImage || bgImage === 'none') return null;

  let resolvedBgImage = bgImage;
  if (style) {
    resolvedBgImage = resolveCssVariables(bgImage, style);
  }

  // 1. Extract content inside function(...)
  // Handles linear-gradient(...), radial-gradient(...), repeating-linear-gradient(...)
  const match = resolvedBgImage.match(/gradient\((.*)\)/);
  if (!match) return null;

  const content = match[1];

  // 2. Split by comma, respecting parentheses (to avoid splitting inside rgb(), oklch(), etc.)
  const parts = [];
  let current = '';
  let parenDepth = 0;

  for (const char of content) {
    if (char === '(') parenDepth++;
    if (char === ')') parenDepth--;
    if (char === ',' && parenDepth === 0) {
      parts.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  if (current) parts.push(current.trim());

  // 3. Find first part that is a color (skip angle/direction)
  for (const part of parts) {
    // Ignore directions (to right) or angles (90deg, 0.5turn)
    if (/^(to\s|[\d.]+(deg|rad|turn|grad))/.test(part)) continue;

    // Extract color: Remove trailing position (e.g. "red 50%" -> "red")
    // Regex matches whitespace + number + unit at end of string
    const colorPart = part.replace(/\s+(-?[\d.]+(%|px|em|rem|ch|vh|vw)?)$/, '');

    // Check if it's not just a number (some gradients might have bare numbers? unlikely in standard syntax)
    if (colorPart) return colorPart;
  }

  return null;
}

function mapDashType(style) {
  if (style === 'dashed') return 'dash';
  if (style === 'dotted') return 'dot';
  return 'solid';
}

/**
 * Analyzes computed border styles and determines the rendering strategy.
 */
export function getBorderInfo(style, scale) {
  const topColorObj = parseColor(style.borderTopColor, style);
  const rightColorObj = parseColor(style.borderRightColor, style);
  const bottomColorObj = parseColor(style.borderBottomColor, style);
  const leftColorObj = parseColor(style.borderLeftColor, style);

  const top = {
    width: parseFloat(style.borderTopWidth) || 0,
    style: style.borderTopStyle,
    color: topColorObj.hex,
    opacity: topColorObj.opacity,
  };
  const right = {
    width: parseFloat(style.borderRightWidth) || 0,
    style: style.borderRightStyle,
    color: rightColorObj.hex,
    opacity: rightColorObj.opacity,
  };
  const bottom = {
    width: parseFloat(style.borderBottomWidth) || 0,
    style: style.borderBottomStyle,
    color: bottomColorObj.hex,
    opacity: bottomColorObj.opacity,
  };
  const left = {
    width: parseFloat(style.borderLeftWidth) || 0,
    style: style.borderLeftStyle,
    color: leftColorObj.hex,
    opacity: leftColorObj.opacity,
  };

  const hasAnyBorder = top.width > 0 || right.width > 0 || bottom.width > 0 || left.width > 0;
  if (!hasAnyBorder) return { type: 'none' };

  // Check if all sides are uniform
  const isUniform =
    top.width === right.width &&
    top.width === bottom.width &&
    top.width === left.width &&
    top.style === right.style &&
    top.style === bottom.style &&
    top.style === left.style &&
    top.color === right.color &&
    top.color === bottom.color &&
    top.color === left.color;

  if (isUniform) {
    return {
      type: 'uniform',
      options: {
        width: top.width * 0.75 * scale,
        color: top.color,
        transparency: (1 - top.opacity) * 100,
        dashType: mapDashType(top.style),
      },
    };
  } else {
    return {
      type: 'composite',
      sides: { top, right, bottom, left },
    };
  }
}

/**
 * Generates an SVG image for composite borders that respects border-radius.
 */
export function generateCompositeBorderSVG(w, h, radius, sides) {
  radius = radius / 2; // Adjust for SVG rendering
  const clipId = 'clip_' + Math.random().toString(36).substr(2, 9);
  let borderRects = '';

  const dashAttributes = (side) => {
    if (side.style === 'dashed') {
      return `stroke-dasharray="${side.width * 3} ${side.width * 2}"`;
    }
    if (side.style === 'dotted') {
      return `stroke-dasharray="0.01 ${side.width * 2}" stroke-linecap="round"`;
    }
    return '';
  };

  const line = (side, x1, y1, x2, y2) =>
    `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" ` +
    `stroke="#${side.color}" stroke-opacity="${side.opacity ?? 1}" ` +
    `stroke-width="${side.width}" ${dashAttributes(side)} />`;

  if (sides.top.width > 0 && sides.top.color) {
    borderRects +=
      sides.top.style === 'dashed' || sides.top.style === 'dotted'
        ? line(sides.top, 0, sides.top.width / 2, w, sides.top.width / 2)
        : `<rect x="0" y="0" width="${w}" height="${sides.top.width}" fill="#${sides.top.color}" fill-opacity="${sides.top.opacity ?? 1}" />`;
  }
  if (sides.right.width > 0 && sides.right.color) {
    borderRects +=
      sides.right.style === 'dashed' || sides.right.style === 'dotted'
        ? line(sides.right, w - sides.right.width / 2, 0, w - sides.right.width / 2, h)
        : `<rect x="${w - sides.right.width}" y="0" width="${sides.right.width}" height="${h}" fill="#${sides.right.color}" fill-opacity="${sides.right.opacity ?? 1}" />`;
  }
  if (sides.bottom.width > 0 && sides.bottom.color) {
    borderRects +=
      sides.bottom.style === 'dashed' || sides.bottom.style === 'dotted'
        ? line(sides.bottom, 0, h - sides.bottom.width / 2, w, h - sides.bottom.width / 2)
        : `<rect x="0" y="${h - sides.bottom.width}" width="${w}" height="${sides.bottom.width}" fill="#${sides.bottom.color}" fill-opacity="${sides.bottom.opacity ?? 1}" />`;
  }
  if (sides.left.width > 0 && sides.left.color) {
    borderRects +=
      sides.left.style === 'dashed' || sides.left.style === 'dotted'
        ? line(sides.left, sides.left.width / 2, 0, sides.left.width / 2, h)
        : `<rect x="0" y="0" width="${sides.left.width}" height="${h}" fill="#${sides.left.color}" fill-opacity="${sides.left.opacity ?? 1}" />`;
  }

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
        <defs>
            <clipPath id="${clipId}">
                <rect x="0" y="0" width="${w}" height="${h}" rx="${radius}" ry="${radius}" />
            </clipPath>
        </defs>
        <g clip-path="url(#${clipId})">
            ${borderRects}
        </g>
    </svg>`;

  return 'data:image/svg+xml;base64,' + btoa(svg.trim());
}

/**
 * Generates an SVG data URL for a solid shape with non-uniform corner radii.
 */
export function generateCustomShapeSVG(w, h, color, opacity, radii, border = null) {
  let { tl, tr, br, bl } = radii;

  // Clamp radii using CSS spec logic (avoid overlap)
  const factor = Math.min(
    w / (tl + tr) || Infinity,
    h / (tr + br) || Infinity,
    w / (br + bl) || Infinity,
    h / (bl + tl) || Infinity
  );

  if (factor < 1) {
    tl *= factor;
    tr *= factor;
    br *= factor;
    bl *= factor;
  }

  const roundedPath = (inset, pathRadii) => {
    const left = inset;
    const top = inset;
    const right = Math.max(left, w - inset);
    const bottom = Math.max(top, h - inset);
    const rtl = Math.max(0, pathRadii.tl);
    const rtr = Math.max(0, pathRadii.tr);
    const rbr = Math.max(0, pathRadii.br);
    const rbl = Math.max(0, pathRadii.bl);
    return `
    M ${left + rtl} ${top}
    L ${right - rtr} ${top}
    A ${rtr} ${rtr} 0 0 1 ${right} ${top + rtr}
    L ${right} ${bottom - rbr}
    A ${rbr} ${rbr} 0 0 1 ${right - rbr} ${bottom}
    L ${left + rbl} ${bottom}
    A ${rbl} ${rbl} 0 0 1 ${left} ${bottom - rbl}
    L ${left} ${top + rtl}
    A ${rtl} ${rtl} 0 0 1 ${left + rtl} ${top}
    Z
  `;
  };

  const path = roundedPath(0, { tl, tr, br, bl });
  const fillPath = color ? `<path d="${path}" fill="#${color}" fill-opacity="${opacity}" />` : '';
  let borderPath = '';
  if (border?.color && border.width > 0) {
    const inset = border.width / 2;
    const strokePath = roundedPath(inset, {
      tl: tl - inset,
      tr: tr - inset,
      br: br - inset,
      bl: bl - inset,
    });
    borderPath = `<path d="${strokePath}" fill="none" stroke="#${border.color}" stroke-width="${border.width}" stroke-opacity="${border.opacity ?? 1}" />`;
  }

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
      ${fillPath}
      ${borderPath}
    </svg>`;

  return 'data:image/svg+xml;base64,' + btoa(svg.trim());
}

// --- REPLACE THE EXISTING parseColor FUNCTION ---
export function resolveCssVariables(value, style) {
  if (!value || typeof value !== 'string') return value;
  let resolved = value;
  let match;
  let iterations = 0;
  // Limit to 10 iterations to prevent infinite loop
  while ((match = resolved.match(/var\((--[a-zA-Z0-9_-]+)\)/)) && iterations < 10) {
    iterations++;
    const varName = match[1];
    const varValue = style.getPropertyValue(varName).trim();
    if (varValue) {
      resolved = resolved.replace(match[0], varValue);
    } else {
      break;
    }
  }
  return resolved;
}

export function parseColor(str, style) {
  if (!str) return { hex: null, opacity: 0 };
  let resolvedStr = str;
  if (style) {
    resolvedStr = resolveCssVariables(str, style);
  }
  if (resolvedStr === 'transparent' || resolvedStr.trim() === 'rgba(0, 0, 0, 0)') {
    return { hex: null, opacity: 0 };
  }

  const ctx = getCtx();
  ctx.fillStyle = str;
  const computed = ctx.fillStyle;

  // 1. Handle Hex Output (e.g. #ff0000) - Fast Path
  if (computed.startsWith('#')) {
    let hex = computed.slice(1);
    let opacity = 1;
    if (hex.length === 3)
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    if (hex.length === 4)
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    if (hex.length === 8) {
      opacity = parseInt(hex.slice(6), 16) / 255;
      hex = hex.slice(0, 6);
    }
    return { hex: hex.toUpperCase(), opacity };
  }

  // 2. Handle RGB/RGBA Output (standard) - Fast Path
  if (computed.startsWith('rgb')) {
    const match = computed.match(/[\d.]+/g);
    if (match && match.length >= 3) {
      const r = parseInt(match[0]);
      const g = parseInt(match[1]);
      const b = parseInt(match[2]);
      const a = match.length > 3 ? parseFloat(match[3]) : 1;
      const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
      return { hex, opacity: a };
    }
  }

  // 3. Fallback: Browser returned a format we don't parse (oklch, lab, color(srgb...), etc.)
  // Use Canvas API to convert to sRGB
  ctx.clearRect(0, 0, 1, 1);
  ctx.fillRect(0, 0, 1, 1);
  const data = ctx.getImageData(0, 0, 1, 1).data;
  // data = [r, g, b, a]
  const r = data[0];
  const g = data[1];
  const b = data[2];
  const a = data[3] / 255;

  if (a === 0) return { hex: null, opacity: 0 };

  const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  return { hex, opacity: a };
}

export function flattenColor(color, node, startFromParent = true) {
  if (!color || !color.hex || color.opacity === 1 || color.opacity === 0) {
    return color;
  }

  let r = parseInt(color.hex.slice(0, 2), 16);
  let g = parseInt(color.hex.slice(2, 4), 16);
  let b = parseInt(color.hex.slice(4, 6), 16);
  let a = color.opacity;

  let current = node;
  if (startFromParent && node) {
    current = node.parentElement;
  }

  while (current && current !== document) {
    const style = window.getComputedStyle(current);
    const bgStr = style.backgroundColor;
    const bg = parseColor(bgStr, style);

    if (bg.hex && bg.opacity > 0) {
      const bgR = parseInt(bg.hex.slice(0, 2), 16);
      const bgG = parseInt(bg.hex.slice(2, 4), 16);
      const bgB = parseInt(bg.hex.slice(4, 6), 16);
      const bgA = bg.opacity;

      const aOut = a + bgA * (1 - a);
      if (aOut > 0) {
        r = (r * a + bgR * bgA * (1 - a)) / aOut;
        g = (g * a + bgG * bgA * (1 - a)) / aOut;
        b = (b * a + bgB * bgA * (1 - a)) / aOut;
        a = aOut;
      }

      if (a >= 1) {
        a = 1;
        break;
      }
    }
    current = current.parentElement;
  }

  if (a < 1) {
    const bgR = 255;
    const bgG = 255;
    const bgB = 255;
    const bgA = 1;

    const aOut = a + bgA * (1 - a);
    if (aOut > 0) {
      r = (r * a + bgR * bgA * (1 - a)) / aOut;
      g = (g * a + bgG * bgA * (1 - a)) / aOut;
      b = (b * a + bgB * bgA * (1 - a)) / aOut;
    }
  }

  const rHex = Math.round(r).toString(16).padStart(2, '0');
  const gHex = Math.round(g).toString(16).padStart(2, '0');
  const bHex = Math.round(b).toString(16).padStart(2, '0');
  const hex = (rHex + gHex + bHex).toUpperCase();

  return { hex, opacity: 1 };
}

export function getPadding(style, scale) {
  const pxToInch = 1 / 96;
  return [
    (parseFloat(style.paddingTop) || 0) * pxToInch * scale,
    (parseFloat(style.paddingRight) || 0) * pxToInch * scale,
    (parseFloat(style.paddingBottom) || 0) * pxToInch * scale,
    (parseFloat(style.paddingLeft) || 0) * pxToInch * scale,
  ];
}

export function getSoftEdges(filterStr, scale) {
  if (!filterStr || filterStr === 'none') return null;
  const match = filterStr.match(/blur\(([\d.]+)px\)/);
  if (match) return parseFloat(match[1]) * 0.75 * scale;
  return null;
}

export function getTextStyle(style, scale, includeMargins = true, inheritedOpacity = 1) {
  let colorObj = parseColor(style.color, style);
  let opacity = colorObj.opacity * inheritedOpacity;

  // Combine text color alpha with element-level opacity
  const elOpacity = parseFloat(style.opacity);
  if (!isNaN(elOpacity)) {
    opacity *= elOpacity;
  }

  const bgClip = style.webkitBackgroundClip || style.backgroundClip;
  if (colorObj.opacity === 0 && bgClip === 'text') {
    const fallback = getGradientFallbackColor(style.backgroundImage, style);
    if (fallback) colorObj = parseColor(fallback, style);
  }

  let lineSpacing = null;
  const fontSizePx = parseFloat(style.fontSize);
  const lhStr = style.lineHeight;

  if (lhStr && lhStr !== 'normal') {
    let lhPx = parseFloat(lhStr);

    // Edge Case: If browser returns a raw multiplier (e.g. "1.5")
    // we must multiply by font size to get the height in pixels.
    // (Note: getComputedStyle usually returns 'px', but inline styles might differ)
    if (/^[0-9.]+$/.test(lhStr)) {
      lhPx = lhPx * fontSizePx;
    }

    if (!isNaN(lhPx) && lhPx > 0) {
      // Convert Pixel Height to Point Height (1px = 0.75pt)
      // And apply the global layout scale.
      lineSpacing = lhPx * 0.75 * scale;
    }
  }

  // --- Spacing (Margins) ---
  // Convert CSS margins (px) to PPTX Paragraph Spacing (pt).
  let paraSpaceBefore = 0;
  let paraSpaceAfter = 0;

  if (includeMargins) {
    const mt = parseFloat(style.marginTop) || 0;
    const mb = parseFloat(style.marginBottom) || 0;

    if (mt > 0) paraSpaceBefore = mt * 0.75 * scale;
    if (mb > 0) paraSpaceAfter = mb * 0.75 * scale;
  }

  const transparency = Math.round((1 - opacity) * 100);

  const textDecoration = `${style.textDecorationLine || ''} ${style.textDecoration || ''}`;
  const verticalAlign = String(style.verticalAlign || '').toLowerCase();

  return {
    color: colorObj.hex || '000000',
    ...(transparency > 0 && { transparency }),
    fontFace: style.fontFamily.split(',')[0].replace(/['"]/g, ''),
    fontSize: fontSizePx * 0.75 * scale,
    bold: parseInt(style.fontWeight) >= 600,
    italic: style.fontStyle === 'italic',
    underline: textDecoration.includes('underline'),
    strike: textDecoration.includes('line-through'),
    ...(verticalAlign === 'super' && { superscript: true }),
    ...(verticalAlign === 'sub' && { subscript: true }),
    // Only add if we have a valid value
    ...(lineSpacing && { lineSpacing }),
    ...(paraSpaceBefore > 0 && { paraSpaceBefore }),
    ...(paraSpaceAfter > 0 && { paraSpaceAfter }),
    // Map background color to highlight if present
    ...(parseColor(style.backgroundColor, style).hex
      ? { highlight: parseColor(style.backgroundColor, style).hex }
      : {}),
    // Mapping letter-spacing to charSpacing
    ...(style.letterSpacing && style.letterSpacing !== 'normal'
      ? { charSpacing: parseFloat(style.letterSpacing) * 0.75 * scale }
      : {}),
  };
}

/**
 * Returns true when the element itself is not painted. `visibility` is
 * intentionally included even though a descendant may restore
 * `visibility:visible`; callers that prune whole subtrees must handle that CSS
 * distinction separately.
 */
export function isVisuallySuppressed(node) {
  if (!node || node.nodeType !== 1) return false;
  const style = window.getComputedStyle(node);
  const opacity = parseFloat(style.opacity);
  return (
    style.display === 'none' ||
    style.visibility === 'hidden' ||
    style.visibility === 'collapse' ||
    (!Number.isNaN(opacity) && opacity <= 0)
  );
}

function hasVisibleDescendantThroughHiddenVisibility(node) {
  for (const child of Array.from(node?.children || [])) {
    const style = window.getComputedStyle(child);
    const opacity = parseFloat(style.opacity);

    // These properties suppress the complete descendant subtree.  Unlike
    // visibility, neither can be restored by a child declaration.
    if (style.display === 'none' || (!Number.isNaN(opacity) && opacity <= 0)) {
      continue;
    }

    if (style.visibility === 'visible') return true;
    if (hasVisibleDescendantThroughHiddenVisibility(child)) return true;
  }
  return false;
}

/**
 * Determines if a given DOM node is primarily a text container.
 * Updated to correctly reject Icon elements so they are rendered as images.
 */
export function isTextContainer(node) {
  const hasText = node.textContent.trim().length > 0;
  if (!hasText) return false;

  const nodeStyle = window.getComputedStyle(node);
  if (nodeStyle.visibility === 'hidden' || nodeStyle.visibility === 'collapse') {
    return false;
  }

  const children = Array.from(node.children).filter((child) => {
    const childStyle = window.getComputedStyle(child);
    const opacity = parseFloat(childStyle.opacity);
    if (childStyle.display === 'none' || (!Number.isNaN(opacity) && opacity <= 0)) {
      return false;
    }
    if (childStyle.visibility === 'hidden' || childStyle.visibility === 'collapse') {
      // A visibility-hidden wrapper can contain a visibility-visible child.
      // Keep such a wrapper in the structural analysis so the ancestor is not
      // collapsed into one text box that would discard the visible override.
      return hasVisibleDescendantThroughHiddenVisibility(child);
    }
    return true;
  });
  if (children.length === 0) return true;

  const isSafeInline = (el) => {
    const tag = (el?.tagName || '').toLowerCase();
    // 1. Reject Web Components / Custom Elements
    if (tag.includes('-')) return false;
    // 2. Reject Explicit Images/SVGs
    if (tag === 'img' || tag === 'svg') return false;

    if (tag === 'i' || tag === 'span') {
      const cls = el.getAttribute('class') || '';
      if (
        typeof cls === 'string' &&
        (cls.includes('fa-') ||
          cls.includes('fas') ||
          cls.includes('far') ||
          cls.includes('fab') ||
          cls.includes('material-icons') ||
          cls.includes('bi-') ||
          cls.includes('icon'))
      ) {
        // Double-check: Must have pseudo-element content to be a CSS icon
        const before = window.getComputedStyle(el, '::before').content;
        const after = window.getComputedStyle(el, '::after').content;
        const hasContent = (c) => c && c !== 'none' && c !== 'normal' && c !== '""';

        if (hasContent(before) || hasContent(after)) return false;
      }
    }

    const style = window.getComputedStyle(el);
    const display = style.display;

    if (style.visibility === 'hidden' || style.visibility === 'collapse') {
      return false;
    }

    // Inline formatting only guarantees shared text flow for true inline
    // boxes. Atomic inline boxes keep their own width/alignment/layout and
    // cannot be represented faithfully as a PowerPoint rich-text run.
    const isAtomicInline = ['inline-block', 'inline-flex', 'inline-grid', 'inline-table'].includes(display);
    if (isAtomicInline) return false;

    // Grid containers need their own render item even as `inline-grid`; folding them into the
    // parent's text run would discard their internal layout and item alignment.
    const isBlockDisplay =
      display === 'block' || display === 'flex' || display.includes('grid') || display === 'table';
    if (isBlockDisplay) return false;

    const parentStyle = el.parentElement ? window.getComputedStyle(el.parentElement) : null;
    const parentDisplay = parentStyle ? parentStyle.display : '';
    const isFlexOrGridItem = parentDisplay.includes('flex') || parentDisplay.includes('grid');
    if (isFlexOrGridItem) return false;

    // 4. Standard Inline Tag Check
    const isInlineTag = ['span', 'b', 'strong', 'em', 'i', 'a', 'small', 'mark'].includes(el.tagName.toLowerCase());
    const isInlineDisplay = display.includes('inline');

    if (!isInlineTag && !isInlineDisplay) return false;

    // 5. Structural Styling Check
    // If a child has a background or border, it's a layout block, not a simple text span.
    const bgColor = parseColor(style.backgroundColor, style);
    const hasVisibleBg = bgColor.hex && bgColor.opacity > 0;
    const hasBorder = parseFloat(style.borderWidth) > 0 && parseColor(style.borderColor, style).opacity > 0;

    if (hasVisibleBg || hasBorder) {
      // Relaxed check: Allow inline elements with background/border to be treated as text.
      // They will be rendered as highlighted text runs (no border support in text runs though).
      // This preserves text flow for "badges".
      // return false;
    }

    // 4. Check for empty shapes (visual objects without text, like dots)
    const hasContent = el.textContent.trim().length > 0;
    if (!hasContent && (hasVisibleBg || hasBorder)) {
      return false;
    }

    return true;
  };

  if (children.every(isSafeInline)) return true;

  // These properties are applied to a PPTX text box, not to an individual
  // rich-text run.  A shared text box is only safe when all direct children
  // agree with their parent on them.
  const normalizeTextAlign = (value) => {
    const align = value || 'start';
    if (align === 'start') return 'left';
    if (align === 'end') return 'right';
    return align;
  };
  const getFlowContext = (style) => ({
    textAlign: normalizeTextAlign(style.textAlign),
    whiteSpace: style.whiteSpace || 'normal',
    direction: style.direction || 'ltr',
    writingMode: style.writingMode || 'horizontal-tb',
  });
  const parentFlowContext = getFlowContext(window.getComputedStyle(node));
  const hasMatchingFlowContext = (el) => {
    const childFlowContext = getFlowContext(window.getComputedStyle(el));
    return Object.entries(parentFlowContext).every(([key, value]) => childFlowContext[key] === value);
  };

  // A common editorial pattern is an inline or heading-style lead followed by
  // one or more normal-flow paragraphs inside a card.  Exporting those children
  // independently gives an inline lead only its natural DOM width.  When
  // PowerPoint wraps it, later paragraph boxes do not move and the text overlaps.
  // Treat only plain, stacked text blocks as one flow; layout containers,
  // decorated blocks and positioned children remain independent PPTX objects.
  const isPlainFlowBlock = (el) => {
    const tag = (el?.tagName || '').toLowerCase();
    if (!['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote'].includes(tag)) {
      return false;
    }

    const style = window.getComputedStyle(el);
    if (style.display !== 'block') return false;
    if (style.position && style.position !== 'static') return false;
    if (style.float && style.float !== 'none') return false;
    if (style.transform && style.transform !== 'none') return false;
    if (style.overflow && style.overflow !== 'visible') return false;
    if (
      parseFloat(style.paddingTop) ||
      parseFloat(style.paddingRight) ||
      parseFloat(style.paddingBottom) ||
      parseFloat(style.paddingLeft)
    ) {
      return false;
    }

    const bgColor = parseColor(style.backgroundColor, style);
    const hasVisibleBg = bgColor.hex && bgColor.opacity > 0;
    const hasBorder = parseFloat(style.borderWidth) > 0 && parseColor(style.borderColor, style).opacity > 0;
    if (hasVisibleBg || hasBorder) return false;

    return Array.from(el.children).every(isSafeInline);
  };

  const hasFlowBlock = children.some(isPlainFlowBlock);
  return (
    hasFlowBlock &&
    children.every((child) => hasMatchingFlowContext(child) && (isSafeInline(child) || isPlainFlowBlock(child)))
  );
}

export function getRotation(transformStr) {
  if (!transformStr || transformStr === 'none') return 0;
  const values = transformStr.split('(')[1].split(')')[0].split(',');
  if (values.length < 4) return 0;
  const a = parseFloat(values[0]);
  const b = parseFloat(values[1]);
  return Math.round(Math.atan2(b, a) * (180 / Math.PI));
}

export function getWritingModeVert(writingMode, textOrientation) {
  const isUpright = textOrientation === 'upright';

  switch (writingMode) {
    case 'vertical-rl':
      return isUpright ? 'wordArtVertRtl' : 'eaVert';
    case 'vertical-lr':
      return isUpright ? 'wordArtVert' : 'mongolianVert';
    case 'sideways-rl':
      return 'vert';
    case 'sideways-lr':
      return 'vert270';
    default:
      return null;
  }
}

export function mapVertToTextDirection(vertVal) {
  if (!vertVal) return null;
  if (vertVal === 'eaVert' || vertVal === 'mongolianVert') return 'vert';
  if (vertVal === 'wordArtVertRtl') return 'wordArtVert';
  if (['vert', 'vert270', 'wordArtVert', 'horz'].includes(vertVal)) return vertVal;
  return null;
}

/**
 * Converts an SVG node to a PNG data URL (rasterized)
 */
export function svgToPng(node) {
  return new Promise((resolve) => {
    const clone = node.cloneNode(true);
    const rect = node.getBoundingClientRect();
    const width = rect.width || 300;
    const height = rect.height || 150;

    inlineSvgStyles(node, clone);
    clone.setAttribute('width', width);
    clone.setAttribute('height', height);
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

    const xml = new XMLSerializer().serializeToString(clone);
    const svgUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(xml)}`;
    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const scale = 3;
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve(null);
    img.src = svgUrl;
  });
}

/**
 * Converts an SVG node to an SVG data URL (preserves vector format)
 * This allows "Convert to Shape" in PowerPoint.
 *
 * @param {SVGElement} node
 * @param {{width?: number, height?: number}} [outputSize] - Optional rendered
 *   dimensions, used when an SVG inside a Shadow Root is sized by its host.
 */
export function svgToSvg(node, outputSize = {}) {
  return new Promise((resolve) => {
    try {
      const clone = node.cloneNode(true);
      const rect = node.getBoundingClientRect();
      const width = outputSize.width || rect.width || 300;
      const height = outputSize.height || rect.height || 150;

      inlineSvgStyles(node, clone);
      clone.setAttribute('width', width);
      clone.setAttribute('height', height);
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

      // Ensure xmlns:xlink is present for any xlink:href attributes
      if (clone.querySelector('[*|href]') || clone.innerHTML.includes('xlink:')) {
        clone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
      }

      const xml = new XMLSerializer().serializeToString(clone);
      // Use base64 encoding for better compatibility with PowerPoint
      const svgUrl = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(xml)))}`;
      resolve(svgUrl);
    } catch (e) {
      console.warn('SVG serialization failed:', e);
      resolve(null);
    }
  });
}

/**
 * Helper to inline computed styles into an SVG clone
 */
function inlineSvgStyles(source, target) {
  const computed = window.getComputedStyle(source);
  const properties = [
    'fill',
    'stroke',
    'stroke-width',
    'stroke-linecap',
    'stroke-linejoin',
    'stroke-opacity',
    'stroke-dasharray',
    'stroke-dashoffset',
    'fill-opacity',
    'fill-rule',
    'clip-rule',
    'stop-color',
    'stop-opacity',
    'vector-effect',
    'paint-order',
    'opacity',
    'font-family',
    'font-size',
    'font-weight',
  ];

  if (computed.fill === 'none') target.setAttribute('fill', 'none');
  else if (computed.fill) target.style.fill = computed.fill;

  if (computed.stroke === 'none') target.setAttribute('stroke', 'none');
  else if (computed.stroke) target.style.stroke = computed.stroke;

  properties.forEach((prop) => {
    if (prop !== 'fill' && prop !== 'stroke') {
      const val = computed.getPropertyValue(prop) || computed[prop];
      if (val && val !== 'auto') target.style[prop] = val;
    }
  });

  for (let i = 0; i < source.children.length; i++) {
    if (target.children[i]) inlineSvgStyles(source.children[i], target.children[i]);
  }
}

export function getVisibleShadow(shadowStr, scale) {
  if (!shadowStr || shadowStr === 'none') return null;
  const shadows = shadowStr.split(/,(?![^()]*\))/);
  for (let s of shadows) {
    s = s.trim();
    if (s.startsWith('rgba(0, 0, 0, 0)')) continue;
    const match = s.match(/(rgba?\([^)]+\)|#[0-9a-fA-F]+)\s+(-?[\d.]+)px\s+(-?[\d.]+)px\s+([\d.]+)px/);
    if (match) {
      const colorStr = match[1];
      const x = parseFloat(match[2]);
      const y = parseFloat(match[3]);
      const blur = parseFloat(match[4]);
      const distance = Math.sqrt(x * x + y * y);
      let angle = Math.atan2(y, x) * (180 / Math.PI);
      if (angle < 0) angle += 360;
      const colorObj = parseColor(colorStr);
      return {
        type: 'outer',
        angle: angle,
        blur: blur * 0.75 * scale,
        offset: distance * 0.75 * scale,
        color: colorObj.hex || '000000',
        opacity: colorObj.opacity,
      };
    }
  }
  return null;
}

/**
 * Generates an SVG image for gradients, supporting degrees and keywords.
 */
export function generateGradientSVG(w, h, bgString, radius, border) {
  try {
    const match = bgString.match(/linear-gradient\((.*)\)/);
    if (!match) return null;
    const content = match[1];

    // Split by comma, ignoring commas inside parentheses (e.g. rgba())
    const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
    if (parts.length < 2) return null;

    let x1 = '50%',
      y1 = '0%',
      x2 = '50%',
      y2 = '100%';
    let stopsStartIndex = 0;
    const firstPart = parts[0].toLowerCase();

    // 1. Check for Keywords (to right, etc.)
    if (firstPart.startsWith('to ')) {
      stopsStartIndex = 1;
      const direction = firstPart.replace('to ', '').trim();
      switch (direction) {
        case 'top':
          x1 = '50%';
          y1 = '100%';
          x2 = '50%';
          y2 = '0%';
          break;
        case 'bottom':
          x1 = '50%';
          y1 = '0%';
          x2 = '50%';
          y2 = '100%';
          break;
        case 'left':
          x1 = '100%';
          y1 = '50%';
          x2 = '0%';
          y2 = '50%';
          break;
        case 'right':
          x1 = '0%';
          y1 = '50%';
          x2 = '100%';
          y2 = '50%';
          break;
        case 'top right':
          x1 = '0%';
          y1 = '100%';
          x2 = '100%';
          y2 = '0%';
          break;
        case 'top left':
          x1 = '100%';
          y1 = '100%';
          x2 = '0%';
          y2 = '0%';
          break;
        case 'bottom right':
          x2 = '100%';
          y2 = '100%';
          break;
        case 'bottom left':
          x1 = '100%';
          y2 = '100%';
          break;
      }
    }
    // 2. Check for Degrees (45deg, 90deg, etc.)
    else if (firstPart.match(/^-?[\d.]+(deg|rad|turn|grad)$/)) {
      stopsStartIndex = 1;
      const val = parseFloat(firstPart);
      // CSS 0deg is Top (North), 90deg is Right (East), 180deg is Bottom (South)
      // We convert this to SVG coordinates on a unit square (0-100%).
      // Formula: Map angle to perimeter coordinates.
      if (!isNaN(val)) {
        let deg = val;
        if (firstPart.endsWith('rad')) deg = val * (180 / Math.PI);
        if (firstPart.endsWith('turn')) deg = val * 360;
        if (firstPart.endsWith('grad')) deg = val * 0.9;

        // CSS angles start at the top and rotate clockwise. SVG coordinates
        // point downwards, hence dy = -cos(theta). Project the box corners on
        // that vector so non-square gradients retain CSS's full coverage.
        const theta = (deg * Math.PI) / 180;
        const dx = Math.sin(theta);
        const dy = -Math.cos(theta);
        const halfLength = (Math.abs(w * dx) + Math.abs(h * dy)) / 2;
        const startX = w / 2 - dx * halfLength;
        const startY = h / 2 - dy * halfLength;
        const endX = w / 2 + dx * halfLength;
        const endY = h / 2 + dy * halfLength;
        const percent = (value, size) => {
          const rounded = Math.round((value / size) * 1000) / 10;
          return `${Number.isInteger(rounded) ? rounded : rounded.toFixed(1)}%`;
        };

        x1 = percent(startX, w);
        y1 = percent(startY, h);
        x2 = percent(endX, w);
        y2 = percent(endY, h);
      }
    }

    // 3. Process Color Stops
    let stopsXML = '';
    const stopParts = parts.slice(stopsStartIndex);

    stopParts.forEach((part, idx) => {
      // Parse "Color Position" (e.g., "red 50%")
      // Regex looks for optional space + number + unit at the end of the string
      let color = part;
      let offset = Math.round((idx / (stopParts.length - 1)) * 100) + '%';

      const posMatch = part.match(/^(.*?)\s+(-?[\d.]+(?:%|px)?)$/);
      if (posMatch) {
        color = posMatch[1];
        offset = posMatch[2];
      }

      // Handle RGBA/RGB for SVG compatibility
      let opacity = 1;
      if (color.includes('rgba')) {
        const rgbaMatch = color.match(/[\d.]+/g);
        if (rgbaMatch && rgbaMatch.length >= 4) {
          opacity = rgbaMatch[3];
          color = `rgb(${rgbaMatch[0]},${rgbaMatch[1]},${rgbaMatch[2]})`;
        }
      }

      stopsXML += `<stop offset="${offset}" stop-color="${color.trim()}" stop-opacity="${opacity}"/>`;
    });

    let strokeAttr = '';
    if (border) {
      strokeAttr = `stroke="#${border.color}" stroke-width="${border.width}"`;
    }

    let tl = 0,
      tr = 0,
      br = 0,
      bl = 0;
    if (typeof radius === 'object' && radius !== null) {
      tl = radius.tl || 0;
      tr = radius.tr || 0;
      br = radius.br || 0;
      bl = radius.bl || 0;
    } else {
      tl = tr = br = bl = radius || 0;
    }

    const factor = Math.min(
      w / (tl + tr) || Infinity,
      h / (tr + br) || Infinity,
      w / (br + bl) || Infinity,
      h / (bl + tl) || Infinity
    );

    if (factor < 1) {
      tl *= factor;
      tr *= factor;
      br *= factor;
      bl *= factor;
    }

    // Generate absolute path based on radius bounds
    const pathD = `M ${tl} 0 L ${w - tr} 0 A ${tr} ${tr} 0 0 1 ${w} ${tr} L ${w} ${h - br} A ${br} ${br} 0 0 1 ${w - br} ${h} L ${bl} ${h} A ${bl} ${bl} 0 0 1 0 ${h - bl} L 0 ${tl} A ${tl} ${tl} 0 0 1 ${tl} 0 Z`;

    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
          <defs>
            <linearGradient id="grad" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">
              ${stopsXML}
            </linearGradient>
          </defs>
          <path d="${pathD}" fill="url(#grad)" ${strokeAttr} />
      </svg>`;

    return 'data:image/svg+xml;base64,' + btoa(svg.trim());
  } catch (e) {
    console.warn('Gradient generation failed:', e);
    return null;
  }
}

export function generateBlurredSVG(w, h, color, radius, blurPx) {
  const padding = blurPx * 3;
  const fullW = w + padding * 2;
  const fullH = h + padding * 2;
  const x = padding;
  const y = padding;
  let shapeTag;
  const isCircle = radius >= Math.min(w, h) / 2 - 1 && Math.abs(w - h) < 2;

  if (isCircle) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    shapeTag = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="#${color}" filter="url(#f1)" />`;
  } else {
    shapeTag = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${radius}" ry="${radius}" fill="#${color}" filter="url(#f1)" />`;
  }

  const svg = `
  <svg xmlns="http://www.w3.org/2000/svg" width="${fullW}" height="${fullH}" viewBox="0 0 ${fullW} ${fullH}">
    <defs>
      <filter id="f1" x="-50%" y="-50%" width="200%" height="200%">
        <feGaussianBlur in="SourceGraphic" stdDeviation="${blurPx}" />
      </filter>
    </defs>
    ${shapeTag}
  </svg>`;

  return {
    data: 'data:image/svg+xml;base64,' + btoa(svg.trim()),
    padding: padding,
  };
}

// src/utils.js

// ... (keep all existing exports) ...

/**
 * Traverses the target DOM and collects all unique font-family names used.
 */
export function getUsedFontFamilies(root) {
  const families = new Set();

  function scan(node) {
    if (node.nodeType === 1) {
      // Element
      const style = window.getComputedStyle(node);
      const fontList = style.fontFamily.split(',');
      // The first font in the stack is the primary one
      const primary = fontList[0].trim().replace(/['"]/g, '');
      if (primary) families.add(primary);
    }
    for (const child of node.childNodes) {
      scan(child);
    }
  }

  // Handle array of roots or single root
  const elements = Array.isArray(root) ? root : [root];
  elements.forEach((el) => {
    const node = typeof el === 'string' ? document.querySelector(el) : el;
    if (node) scan(node);
  });

  return families;
}

// Helper to extract a clean URL from a CSS `src` string. Exported so
// getFontsFromStyleSheets can be unit-tested against synthetic CSSOM.
function extractFontUrl(srcStr) {
  // Look for url("..."), url('...'), or url(...).
  // Prefer woff/ttf/otf; fall back to whatever is available.
  const matches = srcStr.match(/url\((['"]?)(.*?)\1\)/g);
  if (!matches) return null;

  let chosenUrl = null;
  for (const match of matches) {
    const urlRaw = match.replace(/url\((['"]?)(.*?)\1\)/, '$2');
    // Skip data URIs for now (unless we want to support base64 embedding).
    if (urlRaw.startsWith('data:')) continue;

    if (urlRaw.includes('.ttf') || urlRaw.includes('.otf') || urlRaw.includes('.woff')) {
      chosenUrl = urlRaw;
      break;
    }
    if (!chosenUrl) chosenUrl = urlRaw;
  }
  return chosenUrl;
}

/**
 * Extract PowerPoint speaker-notes text from a slide root element.
 *
 * DOM convention: any descendant of the slide root carrying a
 * `data-pptx-notes` attribute contributes its text content to the
 * slide's speaker-notes pane. Elements without the attribute are
 * ignored; elements with the attribute but no text content are ignored
 * (so a stray empty `<template data-pptx-notes>` is harmless).
 *
 * Multiple annotated elements concatenate in document order, separated
 * by a blank line.
 *
 * Three usage patterns work interchangeably:
 *
 *   <template data-pptx-notes>
 *     Speaker notes here. `<template>` content is inert in the DOM so
 *     the notes never render on-slide.
 *   </template>
 *
 *   <div data-pptx-notes hidden>
 *     Also fine. `hidden` keeps the div off-slide visually.
 *   </div>
 *
 *   <p data-pptx-notes style="display: none">
 *     Any element, hidden however you prefer.
 *   </p>
 *
 * If the annotated element is visible (no `hidden`, no CSS `display:
 * none`), its text will appear both on the slide and in the notes pane.
 * That is user error, not a defect of this extractor.
 *
 * @param {Element} root
 * @returns {string} Notes text, trimmed. Empty string if no notes.
 */
export function extractSpeakerNotesFromElement(root) {
  if (!root || typeof root.querySelectorAll !== 'function') return '';
  const nodes = [
    ...(root.matches?.('[data-pptx-notes]') ? [root] : []),
    ...Array.from(root.querySelectorAll('[data-pptx-notes]')),
  ];
  const parts = [];
  for (const node of nodes) {
    const attributeValue = node.getAttribute('data-pptx-notes');
    // <template> stores its markup in a DocumentFragment on `.content`;
    // its own `.textContent` is empty. Fall back to `.textContent` for
    // every other element type.
    const isTemplate = (node?.tagName || '').toLowerCase() === 'template';
    const contentValue = isTemplate && node.content ? node.content.textContent : node.textContent;
    const raw = attributeValue && attributeValue.trim() ? attributeValue : contentValue;
    if (!raw) continue;
    const trimmed = raw.trim();
    if (trimmed) parts.push(trimmed);
  }
  return parts.join('\n\n');
}

/**
 * Scans document.styleSheets to find @font-face URLs for the requested families.
 * Returns an array of { name, url } objects.
 * Walk a list of CSSStyleSheet-like objects and collect @font-face
 * declarations whose family matches usedFamilies.
 *
 * Recurses into @import rules so that a `theme.css` containing
 * `@import url('./fonts.css');` still surfaces the @font-face
 * declarations inside `fonts.css`. Without this, a common CSS
 * organisation (imports for shared type stacks) silently produces an
 * empty embedded-font list.
 *
 * Extracted from getAutoDetectedFonts so it can be exercised in unit
 * tests without depending on document.styleSheets.
 *
 * @param {Set<string>} usedFamilies
 * @param {ArrayLike<CSSStyleSheet>} styleSheets
 * @returns {Array<{name: string, url: string, weight: string, style: string}>}
 */
/**
 * Parses @font-face declarations from raw CSS text string.
 * Used as a fallback when document.styleSheets[i].cssRules is blocked by CORS.
 */
export function parseFontFacesFromCssText(cssText, usedFamilies, baseHref) {
  const foundFonts = [];
  const processedUrls = new Set();

  const fontFaceRegex = /@font-face\s*\{([^}]+)\}/gi;
  let match;
  while ((match = fontFaceRegex.exec(cssText)) !== null) {
    const block = match[1];
    const familyMatch = /font-family\s*:\s*['"]?([^;'"]+)['"]?/i.exec(block);
    if (!familyMatch) continue;
    const familyName = familyMatch[1].trim();

    let matchedFamily = null;
    if (usedFamilies.has(familyName)) {
      matchedFamily = familyName;
    } else {
      for (const fam of usedFamilies) {
        if (fam.toLowerCase() === familyName.toLowerCase()) {
          matchedFamily = fam;
          break;
        }
      }
    }
    if (!matchedFamily) continue;

    const srcMatch = /src\s*:\s*([^;]+)/i.exec(block);
    if (!srcMatch) continue;

    let url = extractFontUrl(srcMatch[1]);
    // Relative URLs in fetched CSS text are relative to the stylesheet href,
    // not the document; resolve them when the caller provides the base.
    if (url && baseHref) {
      try {
        url = new URL(url, baseHref).href;
      } catch (e) {
        // keep original url if it cannot be resolved
      }
    }
    if (url && !processedUrls.has(url)) {
      processedUrls.add(url);
      const weightMatch = /font-weight\s*:\s*([^;]+)/i.exec(block);
      const styleMatch = /font-style\s*:\s*([^;]+)/i.exec(block);
      const weight = weightMatch ? weightMatch[1].trim() : '400';
      const fontStyle = styleMatch ? styleMatch[1].trim().toLowerCase() : 'normal';

      foundFonts.push({ name: matchedFamily, url, weight, style: fontStyle });
    }
  }
  return foundFonts;
}

export function getFontsFromStyleSheets(usedFamilies, styleSheets, blockedHrefs = null) {
  const foundFonts = [];
  const processedUrls = new Set();
  const visitedSheets = new Set(); // Guard against cyclic @import graphs.

  const walk = (sheet) => {
    if (!sheet || visitedSheets.has(sheet)) return;
    visitedSheets.add(sheet);

    let rules;
    try {
      rules = sheet.cssRules || sheet.rules;
    } catch (e) {
      // SecurityError is common for cross-origin sheets (Google Fonts etc.);
      // record sheet.href for asynchronous fetch fallback.
      if (sheet.href && blockedHrefs) {
        blockedHrefs.add(sheet.href);
      }
      console.warn('Cannot scan stylesheet for fonts (CORS restriction):', sheet.href, e && e.message);
      return;
    }
    if (!rules) return;

    for (const rule of Array.from(rules)) {
      // CSSImportRule (type === 3): recurse into the imported sheet so
      // that `@import url('./fonts.css')` in a top-level stylesheet
      // still surfaces its @font-face rules.
      if (rule.constructor.name === 'CSSImportRule' || rule.type === 3) {
        walk(rule.styleSheet);
        continue;
      }

      if (rule.constructor.name === 'CSSFontFaceRule' || rule.type === 5) {
        const familyName = rule.style.getPropertyValue('font-family').replace(/['"]/g, '').trim();

        if (!usedFamilies.has(familyName)) continue;

        const src = rule.style.getPropertyValue('src');
        let url = extractFontUrl(src);

        // Browsers return @font-face src URLs exactly as written in the CSS.
        // For external stylesheets (e.g. assets/deck.css declaring
        // src: url('fonts/x.woff2')) those URLs are relative to the
        // stylesheet, but a later fetch() resolves them against the document
        // base — a silent embed failure and system-font fallback. Resolve
        // against the owning stylesheet's href when it has one; hrefless
        // (inline <style>) sheets already resolve correctly at fetch time.
        if (url) {
          const base = (rule.parentStyleSheet && rule.parentStyleSheet.href) || sheet.href;
          if (base) {
            try {
              url = new URL(url, base).href;
            } catch (e) {
              // keep original url if it cannot be resolved
            }
          }
        }

        if (url && !processedUrls.has(url)) {
          processedUrls.add(url);
          // Preserve font-weight and font-style so downstream grouping can
          // classify each @font-face declaration into one of PowerPoint's
          // four embedded-font slots (regular / bold / italic / boldItalic).
          const weight = rule.style.getPropertyValue('font-weight').trim() || '400';
          const fontStyle = (rule.style.getPropertyValue('font-style').trim() || 'normal').toLowerCase();
          foundFonts.push({ name: familyName, url: url, weight: weight, style: fontStyle });
        }
      }
    }
  };

  for (const sheet of Array.from(styleSheets || [])) {
    walk(sheet);
  }

  return foundFonts;
}

/**
 * Scans document.styleSheets and document links to find @font-face URLs for the requested
 * families. Returns an array of { name, url, weight, style } objects.
 * Includes asynchronous fetch fallback for cross-origin stylesheets (e.g. Google Fonts
 * included without crossorigin="anonymous").
 */
export async function getAutoDetectedFonts(usedFamilies) {
  const blockedHrefs = new Set();
  const fontEntries = getFontsFromStyleSheets(usedFamilies, document.styleSheets, blockedHrefs);

  if (typeof document !== 'undefined') {
    const links = document.querySelectorAll('link[rel="stylesheet"]');
    links.forEach((link) => {
      if (link.href && (link.href.includes('fonts.googleapis.com') || link.href.includes('fonts.gstatic.com'))) {
        blockedHrefs.add(link.href);
      }
    });
  }

  if (blockedHrefs.size > 0 && typeof fetch !== 'undefined') {
    const fetchPromises = Array.from(blockedHrefs).map(async (href) => {
      try {
        const res = await fetch(href);
        if (!res.ok) return [];
        const cssText = await res.text();
        return parseFontFacesFromCssText(cssText, usedFamilies, href);
      } catch (e) {
        console.warn('Failed to fetch cross-origin stylesheet fallback for fonts:', href, e);
        return [];
      }
    });

    const fetchedFontLists = await Promise.all(fetchPromises);
    const existingUrls = new Set(fontEntries.map((f) => f.url));

    for (const fontList of fetchedFontLists) {
      for (const f of fontList) {
        if (!existingUrls.has(f.url)) {
          existingUrls.add(f.url);
          fontEntries.push(f);
        }
      }
    }
  }

  return fontEntries;
}

/**
 * Map a CSS font-weight / font-style pair into one of the four slots
 * that PowerPoint's embedded-font list supports: regular, bold, italic,
 * boldItalic. Anything at weight >= 600 counts as bold; italic/oblique
 * counts as italic.
 *
 * Exported so callers can classify a family's variants ahead of grouping,
 * and so the classification is unit-testable.
 */
export function classifyFontVariant(weight, style) {
  const wRaw = String(weight || '400')
    .toLowerCase()
    .trim();
  let w = parseInt(wRaw, 10);
  if (isNaN(w)) {
    if (wRaw === 'bold' || wRaw === 'bolder') w = 700;
    else if (wRaw === 'lighter') w = 300;
    else w = 400;
  }
  const isBold = w >= 600;
  const sRaw = String(style || 'normal').toLowerCase();
  const isItalic = sRaw === 'italic' || sRaw === 'oblique';
  if (isBold && isItalic) return 'boldItalic';
  if (isBold) return 'bold';
  if (isItalic) return 'italic';
  return 'regular';
}

/**
 * Detect when multiple @font-face declarations for the same family map
 * into the same PowerPoint slot with materially different weights.
 *
 * PowerPoint's embedded-font model only exposes four slots per family
 * (regular / bold / italic / boldItalic). CSS weight 700 (Bold) and
 * weight 900 (Black) both classify as `bold`; the second one silently
 * loses glyphs during embed. This is a real workflow surprise for anyone
 * declaring `font-family: 'Inter'; font-weight: 900` alongside 700.
 *
 * Returns an array of { family, variant, weights } collision descriptors
 * so callers can emit an actionable warning.
 *
 * @param {Array<{name: string, weight?: string|number, style?: string}>} fontEntries
 */
export function detectVariantSlotCollisions(fontEntries) {
  const seen = new Map(); // key = "family::variant" -> Set of weight strings
  for (const f of fontEntries) {
    const variant = classifyFontVariant(f.weight, f.style);
    const key = f.name + '::' + variant;
    if (!seen.has(key)) seen.set(key, new Set());
    seen.get(key).add(String(f.weight ?? '400').trim());
  }
  const collisions = [];
  for (const [key, weights] of seen) {
    if (weights.size < 2) continue;
    const [family, variant] = key.split('::');
    collisions.push({ family, variant, weights: Array.from(weights).sort() });
  }
  return collisions;
}

/**
 * Split a text node's value into line segments for an element whose effective
 * CSS `white-space` preserves author line breaks (`pre`, `pre-wrap`, `pre-line`).
 *
 * Returns `[{ text, breakLine }]` where `breakLine` marks a hard line break after
 * that segment (PptxGenJS `breakLine`). Pure string logic — no DOM/canvas — so it
 * is unit-testable in isolation.
 *
 * Rules (per CSS Text spec):
 *  - `pre` / `pre-wrap`: keep newlines AND runs of spaces; tabs become spaces
 *    (PPTX text runs have no tab stops).
 *  - `pre-line`: keep newlines but collapse runs of spaces/tabs.
 *  - A single newline right after a `<pre>` start tag is ignored (HTML parsing).
 *  - A single trailing newline on the last text node is the line terminator and
 *    is dropped, so `<pre>a\n</pre>` is one line, not one line + a blank one.
 */
export function splitPreformattedText(value, whiteSpace, options = {}) {
  const { isFirstChild = false, isLastChild = false, isPre = false, textTransform = 'none' } = options;

  let raw = String(value).replace(/\r\n?/g, '\n');
  if (isFirstChild && isPre && raw[0] === '\n') raw = raw.slice(1);
  raw = whiteSpace === 'pre-line' ? raw.replace(/[ \t]+/g, ' ') : raw.replace(/\t/g, '    ');
  if (isLastChild) raw = raw.replace(/\n$/, '');
  if (!raw.length) return [];

  const transform = (s) => {
    if (textTransform === 'uppercase') return s.toUpperCase();
    if (textTransform === 'lowercase') return s.toLowerCase();
    if (textTransform === 'capitalize') return s.replace(/\b\w/g, (c) => c.toUpperCase());
    return s;
  };

  const lines = raw.split('\n');
  return lines.map((line, i) => ({ text: transform(line), breakLine: i < lines.length - 1 }));
}

/**
 * Resolve the nearest anchor exactly as the browser does. `HTMLAnchorElement.href`
 * expands relative URLs against `<base>`/document URL, unlike getAttribute().
 */
export function getNodeHyperlink(node) {
  if (!node || node.nodeType !== 1 || typeof node.closest !== 'function') return null;
  const anchor = node.closest('a[href]');
  if (!anchor) return null;
  const url = anchor.href || anchor.getAttribute('href');
  if (!url) return null;
  return {
    url,
    tooltip: anchor.getAttribute('title') || undefined,
  };
}

function isBlockFlowDisplay(display) {
  const normalized = String(display || '').toLowerCase();
  return (
    ['block', 'list-item', 'flow-root', 'flex', 'grid', 'table'].includes(normalized) || normalized.startsWith('table-')
  );
}

function hasNonZeroBoxSpacing(style) {
  return [
    style.marginTop,
    style.marginRight,
    style.marginBottom,
    style.marginLeft,
    style.paddingTop,
    style.paddingRight,
    style.paddingBottom,
    style.paddingLeft,
  ].some((value) => (parseFloat(value) || 0) !== 0);
}

/**
 * True when textual pseudo-content participates in the host's ordinary inline
 * text flow. Such content belongs in the host rich-text runs; positioned or
 * decorated pseudo-elements remain independent PowerPoint objects.
 */
export function isInlineTextPseudoStyle(style) {
  if (!style) return false;

  const display = String(style.display || '').toLowerCase();
  const position = String(style.position || 'static').toLowerCase();
  const float = String(style.float || 'none').toLowerCase();
  const transform = String(style.transform || 'none').toLowerCase();
  if (display !== 'inline' || position !== 'static' || float !== 'none' || transform !== 'none') {
    return false;
  }

  const background = parseColor(style.backgroundColor, style);
  const border = parseColor(style.borderColor, style);
  const hasBackground =
    (background.hex && background.opacity > 0) || (style.backgroundImage && style.backgroundImage !== 'none');
  const hasBorder = (parseFloat(style.borderWidth) || 0) > 0 && border.opacity > 0;

  return !hasBackground && !hasBorder && !hasNonZeroBoxSpacing(style);
}

function normalizeCollapsibleTextPartBoundaries(parts) {
  let firstInSegment = true;
  let previousTextPart = null;

  for (const part of parts) {
    if (part.options?.breakLine) {
      if (previousTextPart) {
        previousTextPart.text = previousTextPart.text.replace(/[\t\n\f\r ]+$/, '');
      }
      firstInSegment = true;
      previousTextPart = null;
      continue;
    }

    if (typeof part.text !== 'string' || part.text.length === 0) continue;

    if (firstInSegment) {
      part.text = part.text.replace(/^[\t\n\f\r ]+/, '');
    } else if (previousTextPart && /[\t\n\f\r ]$/.test(previousTextPart.text) && /^[\t\n\f\r ]/.test(part.text)) {
      part.text = part.text.replace(/^[\t\n\f\r ]+/, '');
    }

    if (part.text.length > 0) {
      firstInSegment = false;
      previousTextPart = part;
    }
  }

  if (previousTextPart) {
    previousTextPart.text = previousTextPart.text.replace(/[\t\n\f\r ]+$/, '');
  }
}

function parseCounterDirectives(value, defaultValue) {
  const normalized = String(value || '').trim();
  if (!normalized || normalized === 'none') return [];

  const tokens = normalized.split(/\s+/).filter(Boolean);
  const directives = [];
  for (let index = 0; index < tokens.length; index++) {
    const name = tokens[index];
    if (!/^[-_a-z][-_a-z0-9]*$/i.test(name)) continue;
    const following = tokens[index + 1];
    const hasExplicitValue = following !== undefined && /^-?\d+$/.test(following);
    directives.push({ name, value: hasExplicitValue ? Number.parseInt(following, 10) : defaultValue });
    if (hasExplicitValue) index++;
  }
  return directives;
}

function decodeCssString(value) {
  return value
    .replace(/\\([0-9a-f]{1,6})\s?/gi, (_match, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/\\([\\"'])/g, '$1');
}

function alphaCounter(value, upper) {
  if (value <= 0) return String(value);
  let remaining = value;
  let result = '';
  while (remaining > 0) {
    remaining--;
    result = String.fromCharCode(97 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }
  return upper ? result.toUpperCase() : result;
}

function romanCounter(value, upper) {
  if (value <= 0 || value >= 4000) return String(value);
  const numerals = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
    [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
  ];
  let remaining = value;
  let result = '';
  for (const [amount, glyph] of numerals) {
    while (remaining >= amount) {
      result += glyph;
      remaining -= amount;
    }
  }
  return upper ? result.toUpperCase() : result;
}

function formatCounter(value, style = 'decimal') {
  const normalizedStyle = String(style || 'decimal').trim().toLowerCase();
  if (normalizedStyle === 'decimal-leading-zero') {
    return value >= 0 && value < 10 ? `0${value}` : String(value);
  }
  if (normalizedStyle === 'lower-alpha' || normalizedStyle === 'lower-latin') return alphaCounter(value, false);
  if (normalizedStyle === 'upper-alpha' || normalizedStyle === 'upper-latin') return alphaCounter(value, true);
  if (normalizedStyle === 'lower-roman') return romanCounter(value, false);
  if (normalizedStyle === 'upper-roman') return romanCounter(value, true);
  return String(value);
}

function unquoteCssString(value) {
  const normalized = String(value || '').trim();
  const quote = normalized[0];
  if ((quote === '"' || quote === "'") && normalized.at(-1) === quote) {
    return decodeCssString(normalized.slice(1, -1));
  }
  return decodeCssString(normalized);
}

function resolveGeneratedContent(value, counterStacks, node) {
  let normalized = String(value || '').trim();
  if (!normalized || normalized === 'none' || normalized === 'normal' || normalized === '""' || normalized === "''") {
    return '';
  }

  normalized = normalized.replace(
    /counters\(\s*([-_a-z][-_a-z0-9]*)\s*,\s*((?:"(?:\\.|[^"\\])*")|(?:'(?:\\.|[^'\\])*'))(?:\s*,\s*([-_a-z][-_a-z0-9]*))?\s*\)/gi,
    (_match, name, separator, style) => {
      const values = counterStacks.get(name) || [0];
      return values.map((entry) => formatCounter(entry, style)).join(unquoteCssString(separator));
    }
  );
  normalized = normalized.replace(
    /counter\(\s*([-_a-z][-_a-z0-9]*)(?:\s*,\s*([-_a-z][-_a-z0-9]*))?\s*\)/gi,
    (_match, name, style) => formatCounter(counterStacks.get(name)?.at(-1) ?? 0, style)
  );
  normalized = normalized.replace(
    /attr\(\s*([-_a-z][-_a-z0-9]*)\s*\)/gi,
    (_match, name) => node.getAttribute(name) || ''
  );

  const tokens = Array.from(normalized.matchAll(/"((?:\\.|[^"\\])*)"|'((?:\\.|[^'\\])*)'|([^\s]+)/g));
  return tokens
    .map((token) => (token[1] ?? token[2]) === undefined ? token[3] : decodeCssString(token[1] ?? token[2]))
    .join('');
}

function applyCounterStyle(style, counterStacks) {
  const savedScopes = new Map();
  const save = (name) => {
    if (!savedScopes.has(name)) savedScopes.set(name, counterStacks.has(name) ? [...counterStacks.get(name)] : null);
  };

  for (const { name, value } of parseCounterDirectives(style.counterReset, 0)) {
    save(name);
    const stack = counterStacks.get(name) || [];
    counterStacks.set(name, [...stack, value]);
  }
  for (const { name, value } of parseCounterDirectives(style.counterSet, 0)) {
    if (!counterStacks.has(name)) save(name);
    const stack = counterStacks.get(name) || [];
    if (stack.length === 0) stack.push(value);
    else stack[stack.length - 1] = value;
    counterStacks.set(name, stack);
  }
  for (const { name, value } of parseCounterDirectives(style.counterIncrement, 1)) {
    const stack = counterStacks.get(name) || [0];
    stack[stack.length - 1] += value;
    counterStacks.set(name, stack);
  }

  return () => {
    for (const [name, previous] of savedScopes) {
      if (previous === null) counterStacks.delete(name);
      else counterStacks.set(name, previous);
    }
  };
}

/**
 * Resolves generated textual content once in rendered DOM order. Chromium's
 * getComputedStyle() preserves counter()/counters() expressions, so consumers
 * must carry CSS counter state instead of exporting those expressions verbatim.
 */
export function buildPseudoContentMap(root) {
  const resolved = new WeakMap();
  const counters = new Map();

  const resolvePseudo = (node, pseudoType) => {
    const style = window.getComputedStyle(node, pseudoType);
    const rawContent = String(style.content || '').trim();
    const isGenerated =
      style.display !== 'none' && rawContent && rawContent !== 'none' && rawContent !== 'normal';
    if (!isGenerated) {
      const entry = resolved.get(node) || {};
      entry[pseudoType] = '';
      resolved.set(node, entry);
      return;
    }
    const restore = applyCounterStyle(style, counters);
    const content = resolveGeneratedContent(rawContent, counters, node);
    restore();
    const entry = resolved.get(node) || {};
    entry[pseudoType] = content;
    resolved.set(node, entry);
  };

  const walk = (node) => {
    if (!node || node.nodeType !== 1) return;
    const style = window.getComputedStyle(node);
    if (style.display === 'none') return;
    const restore = applyCounterStyle(style, counters);
    resolvePseudo(node, '::before');
    for (const child of node.children) walk(child);
    resolvePseudo(node, '::after');
    restore();
  };

  walk(root);
  return resolved;
}

export function getResolvedPseudoContent(node, pseudoType, pseudoContentByNode = null) {
  const resolved = pseudoContentByNode?.get(node);
  if (resolved && Object.prototype.hasOwnProperty.call(resolved, pseudoType)) return resolved[pseudoType];
  return resolveGeneratedContent(window.getComputedStyle(node, pseudoType).content, new Map(), node);
}

export function collectTextParts(
  node,
  parentStyle,
  scale,
  activeHyperlink = null,
  isRoot = true,
  inheritedOpacity = 1,
  pseudoContentByNode = null
) {
  if (node.nodeType === 1 && isVisuallySuppressed(node)) return [];

  const parts = [];
  let hyperlink = activeHyperlink;

  // Hyperlink inheritance: If no hyperlink is active, check if this node is an <a> or inside one.
  if (!hyperlink && node.nodeType === 1) {
    hyperlink = getNodeHyperlink(node);
  }

  // Check for CSS Content (::before) - often used for icons
  if (node.nodeType === 1) {
    const beforeStyle = window.getComputedStyle(node, '::before');
    const content = getResolvedPseudoContent(node, '::before', pseudoContentByNode);
    if (
      content &&
      content !== 'none' &&
      content !== 'normal' &&
      content !== '""' &&
      isInlineTextPseudoStyle(beforeStyle)
    ) {
      const cleanContent = content;
      if (cleanContent.trim()) {
        const textOpts = getTextStyle(beforeStyle, scale, false, inheritedOpacity);
        if (hyperlink) textOpts.hyperlink = hyperlink;

        // Apply __spc_ suffix if charSpacing is defined
        if (textOpts.charSpacing !== undefined) {
          const spcVal = Math.round(textOpts.charSpacing * 100);
          if (textOpts.fontFace) {
            textOpts.fontFace = `${textOpts.fontFace}__spc_${spcVal}`;
          }
        }

        parts.push({
          text: cleanContent,
          options: textOpts,
        });
      }
    }
  }

  let trimNextLeading = false;

  node.childNodes.forEach((child, index) => {
    if (child.nodeType === 3) {
      // Honor CSS white-space: pre / pre-wrap / pre-line preserve author line
      // breaks (and, except pre-line, runs of spaces). Without this, every newline
      // and indent inside a <pre> / white-space:pre(-wrap) block is collapsed to a
      // single space and multi-line content renders as one run.
      const wsStyle = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
      const whiteSpace = wsStyle.whiteSpace || 'normal';
      if (whiteSpace === 'pre' || whiteSpace === 'pre-wrap' || whiteSpace === 'pre-line') {
        const segs = splitPreformattedText(child.nodeValue, whiteSpace, {
          isFirstChild: index === 0,
          isLastChild: index === node.childNodes.length - 1,
          isPre: node.nodeType === 1 && (node.tagName || '').toLowerCase() === 'pre',
          textTransform: wsStyle.textTransform,
        });
        if (segs.length) {
          const baseOpts = getTextStyle(wsStyle, scale, !isRoot, inheritedOpacity);
          if (hyperlink) baseOpts.hyperlink = hyperlink;
          if (baseOpts.charSpacing !== undefined && baseOpts.fontFace) {
            baseOpts.fontFace = `${baseOpts.fontFace}__spc_${Math.round(baseOpts.charSpacing * 100)}`;
          }
          // Naked text node: don't paint the parent background as a text highlight if it's block-level.
          const display = wsStyle.display || '';
          const isBlock = ['block', 'flex', 'grid', 'table', 'list-item'].includes(String(display).toLowerCase());
          if (isBlock) {
            delete baseOpts.highlight;
          }
          segs.forEach((seg) => {
            parts.push({
              text: seg.text,
              options: seg.breakLine ? { ...baseOpts, breakLine: true } : { ...baseOpts },
            });
          });
        }
        trimNextLeading = false;
        return;
      }

      // Text (white-space: normal / nowrap — collapse runs of whitespace)
      let val = child.nodeValue.replace(/[\t\n\f\r ]+/g, ' ');

      if (isRoot && index === 0) val = val.replace(/^[\t\n\f\r ]+/, '');
      if (trimNextLeading) {
        val = val.replace(/^[\t\n\f\r ]+/, '');
        trimNextLeading = false;
      }
      if (isRoot && index === node.childNodes.length - 1) {
        val = val.replace(/[\t\n\f\r ]+$/, '');
      }

      if (val) {
        // Use parent style if child is text node, otherwise current style
        const styleToUse = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
        const transform = styleToUse.textTransform;
        if (transform === 'uppercase') val = val.toUpperCase();
        else if (transform === 'lowercase') val = val.toLowerCase();
        else if (transform === 'capitalize') val = val.replace(/\b\w/g, (c) => c.toUpperCase());

        const textOpts = getTextStyle(styleToUse, scale, !isRoot, inheritedOpacity);
        if (hyperlink) textOpts.hyperlink = hyperlink;

        // Apply __spc_ suffix if charSpacing is defined
        if (textOpts.charSpacing !== undefined) {
          const spcVal = Math.round(textOpts.charSpacing * 100);
          if (textOpts.fontFace) {
            textOpts.fontFace = `${textOpts.fontFace}__spc_${spcVal}`;
          }
        }

        // BUG FIX: Avoid rendering the parent's background as a text highlight for naked text nodes.
        // The parent container's background is typically already rendered as a Shape Fill.
        // Only delete the highlight if the parent display is block-level.
        if (child.nodeType === 3 && textOpts.highlight) {
          const display = styleToUse.display || '';
          const isBlock = ['block', 'flex', 'grid', 'table', 'list-item'].includes(String(display).toLowerCase());
          if (isBlock) {
            delete textOpts.highlight;
          }
        }

        parts.push({
          text: val,
          options: textOpts,
        });
      }
    } else if (child.nodeType === 1) {
      if (isVisuallySuppressed(child)) return;

      if ((child?.tagName || '').toLowerCase() === 'br') {
        if (parts.length > 0) {
          const lastPart = parts[parts.length - 1];
          if (lastPart.text && typeof lastPart.text === 'string') {
            lastPart.text = lastPart.text.trimEnd();
          }
        }
        parts.push({ text: '', options: { breakLine: true } });
        trimNextLeading = true;
      } else {
        const childStyle = window.getComputedStyle(child);
        const isBlock = isBlockFlowDisplay(childStyle.display);
        if (isBlock && parts.length > 0 && !parts[parts.length - 1].options?.breakLine) {
          parts.push({ text: '', options: { breakLine: true } });
        }

        const nodeStyle = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
        const nodeOpacity = parseFloat(nodeStyle?.opacity);
        const childInheritedOpacity =
          inheritedOpacity * (Number.isNaN(nodeOpacity) ? 1 : Math.max(0, Math.min(1, nodeOpacity)));
        const childParts = collectTextParts(
          child,
          parentStyle,
          scale,
          hyperlink,
          false,
          childInheritedOpacity,
          pseudoContentByNode
        );
        if (childParts.length > 0) parts.push(...childParts);

        if (isBlock) {
          parts.push({ text: '', options: { breakLine: true } });
          trimNextLeading = true;
        }
      }
    }
  });

  // Check for CSS Content (::after) - often used for icons
  if (node.nodeType === 1) {
    const afterStyle = window.getComputedStyle(node, '::after');
    const content = getResolvedPseudoContent(node, '::after', pseudoContentByNode);
    if (
      content &&
      content !== 'none' &&
      content !== 'normal' &&
      content !== '""' &&
      isInlineTextPseudoStyle(afterStyle)
    ) {
      const cleanContent = content;
      if (cleanContent.trim()) {
        const textOpts = getTextStyle(afterStyle, scale, false, inheritedOpacity);
        if (hyperlink) textOpts.hyperlink = hyperlink;

        // Apply __spc_ suffix if charSpacing is defined
        if (textOpts.charSpacing !== undefined) {
          const spcVal = Math.round(textOpts.charSpacing * 100);
          if (textOpts.fontFace) {
            textOpts.fontFace = `${textOpts.fontFace}__spc_${spcVal}`;
          }
        }

        parts.push({
          text: cleanContent,
          options: textOpts,
        });
      }
    }
  }

  // Cleanup potential trailing empty breakLines
  while (parts.length > 0 && parts[parts.length - 1].options?.breakLine && parts[parts.length - 1].text === '') {
    parts.pop();
  }

  if (isRoot) {
    const rootStyle = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
    const whiteSpace = rootStyle?.whiteSpace || 'normal';
    if (whiteSpace === 'normal' || whiteSpace === 'nowrap') {
      normalizeCollapsibleTextPartBoundaries(parts);
    }
  }

  return parts;
}
