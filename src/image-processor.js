// src/image-processor.js

function radiusPair(value) {
  if (typeof value === 'number') return { x: value, y: value };
  if (value && typeof value === 'object') {
    return {
      x: Number.isFinite(value.x) ? value.x : 0,
      y: Number.isFinite(value.y) ? value.y : Number.isFinite(value.x) ? value.x : 0,
    };
  }
  return { x: 0, y: 0 };
}

export function normalizeImageCornerRadii(radius, targetW, targetH) {
  const uniform = typeof radius === 'number' ? radius : null;
  const input = radius && typeof radius === 'object' ? radius : {};
  const radii = {
    tl: radiusPair(uniform ?? input.tl),
    tr: radiusPair(uniform ?? input.tr),
    br: radiusPair(uniform ?? input.br),
    bl: radiusPair(uniform ?? input.bl),
  };
  const factor = Math.min(
    1,
    targetW / (radii.tl.x + radii.tr.x) || 1,
    targetW / (radii.bl.x + radii.br.x) || 1,
    targetH / (radii.tl.y + radii.bl.y) || 1,
    targetH / (radii.tr.y + radii.br.y) || 1
  );
  if (factor < 1) {
    Object.values(radii).forEach((corner) => {
      corner.x *= factor;
      corner.y *= factor;
    });
  }
  return radii;
}

function splitCssPosition(value) {
  const parts = [];
  let depth = 0;
  let current = '';
  for (const character of String(value || '').trim()) {
    if (character === '(') depth += 1;
    if (character === ')') depth = Math.max(0, depth - 1);
    if (/\s/.test(character) && depth === 0) {
      if (current) parts.push(current);
      current = '';
    } else {
      current += character;
    }
  }
  if (current) parts.push(current);
  return parts;
}

function resolvePositionComponent(value, freeSpace) {
  const token = String(value || '').trim().toLowerCase();
  if (token === 'left' || token === 'top') return 0;
  if (token === 'center') return freeSpace / 2;
  if (token === 'right' || token === 'bottom') return freeSpace;
  if (/^-?[\d.]+%$/.test(token)) return freeSpace * (parseFloat(token) / 100);
  if (/^-?[\d.]+px$/.test(token)) return parseFloat(token);

  const calc = token.match(/^calc\(\s*(-?[\d.]+)%\s*([+-])\s*([\d.]+)px\s*\)$/);
  if (calc) {
    const percentageOffset = freeSpace * (parseFloat(calc[1]) / 100);
    const pixelOffset = parseFloat(calc[3]) * (calc[2] === '-' ? -1 : 1);
    return percentageOffset + pixelOffset;
  }
  return freeSpace / 2;
}

export function resolveObjectPositionOffset(objectPosition, targetW, targetH, renderW, renderH) {
  const parts = splitCssPosition(objectPosition);
  const horizontalKeywords = new Set(['left', 'right']);
  const verticalKeywords = new Set(['top', 'bottom']);
  const freeX = targetW - renderW;
  const freeY = targetH - renderH;

  if (parts.length === 4) {
    const offsets = { x: freeX / 2, y: freeY / 2 };
    for (let index = 0; index < parts.length; index += 2) {
      const edge = parts[index];
      const distance = parseFloat(parts[index + 1]) || 0;
      if (horizontalKeywords.has(edge)) offsets.x = edge === 'left' ? distance : freeX - distance;
      if (verticalKeywords.has(edge)) offsets.y = edge === 'top' ? distance : freeY - distance;
    }
    return offsets;
  }

  let xToken = parts[0] || '50%';
  let yToken = parts[1] || '50%';
  if (parts.length === 1 && verticalKeywords.has(xToken)) {
    yToken = xToken;
    xToken = '50%';
  } else if (parts.length >= 2 && verticalKeywords.has(xToken) && horizontalKeywords.has(yToken)) {
    [xToken, yToken] = [yToken, xToken];
  }
  return {
    x: resolvePositionComponent(xToken, freeX),
    y: resolvePositionComponent(yToken, freeY),
  };
}

function traceRoundedRect(ctx, targetW, targetH, radii) {
  const corner = (cx, cy, radius, start, end) => {
    if (radius.x <= 0 || radius.y <= 0) ctx.lineTo(cx, cy);
    else ctx.ellipse(cx, cy, radius.x, radius.y, 0, start, end);
  };

  ctx.moveTo(radii.tl.x, 0);
  ctx.lineTo(targetW - radii.tr.x, 0);
  corner(targetW - radii.tr.x, radii.tr.y, radii.tr, -Math.PI / 2, 0);
  ctx.lineTo(targetW, targetH - radii.br.y);
  corner(targetW - radii.br.x, targetH - radii.br.y, radii.br, 0, Math.PI / 2);
  ctx.lineTo(radii.bl.x, targetH);
  corner(radii.bl.x, targetH - radii.bl.y, radii.bl, Math.PI / 2, Math.PI);
  ctx.lineTo(0, radii.tl.y);
  corner(radii.tl.x, radii.tl.y, radii.tl, Math.PI, (3 * Math.PI) / 2);
  ctx.closePath();
}

export async function getProcessedImage(src, targetW, targetH, radius, objectFit = 'fill', objectPosition = '50% 50%') {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'Anonymous';

    img.onload = () => {
      const canvas = document.createElement('canvas');
      const scale = 2; // Double resolution
      canvas.width = targetW * scale;
      canvas.height = targetH * scale;
      const ctx = canvas.getContext('2d');
      ctx.scale(scale, scale);

      const r = normalizeImageCornerRadii(radius, targetW, targetH);

      // 1. Draw Mask
      ctx.beginPath();
      traceRoundedRect(ctx, targetW, targetH, r);
      ctx.fillStyle = '#000';
      ctx.fill();

      // 2. Composite Source-In
      ctx.globalCompositeOperation = 'source-in';

      // 3. Draw Image with Object Fit logic
      const wRatio = targetW / img.width;
      const hRatio = targetH / img.height;
      let renderW, renderH;

      if (objectFit === 'contain') {
        const fitScale = Math.min(wRatio, hRatio);
        renderW = img.width * fitScale;
        renderH = img.height * fitScale;
      } else if (objectFit === 'cover') {
        const coverScale = Math.max(wRatio, hRatio);
        renderW = img.width * coverScale;
        renderH = img.height * coverScale;
      } else if (objectFit === 'none') {
        renderW = img.width;
        renderH = img.height;
      } else if (objectFit === 'scale-down') {
        const scaleDown = Math.min(1, Math.min(wRatio, hRatio));
        renderW = img.width * scaleDown;
        renderH = img.height * scaleDown;
      } else {
        // 'fill' (default)
        renderW = targetW;
        renderH = targetH;
      }

      const { x: renderX, y: renderY } = resolveObjectPositionOffset(
        objectPosition,
        targetW,
        targetH,
        renderW,
        renderH
      );

      ctx.drawImage(img, renderX, renderY, renderW, renderH);

      resolve(canvas.toDataURL('image/png'));
    };

    img.onerror = () => resolve(null);
    img.src = src;
  });
}
