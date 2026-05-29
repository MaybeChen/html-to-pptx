// src/utils.js
// canvas context for color normalization

let _ctx;

function getCtx() {
  if (!_ctx) _ctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true });
  return _ctx;
}

function getTableBorder(style, side, scale) {
  const widthStr = style[`border${side}Width`];
  const styleStr = style[`border${side}Style`];
  const colorStr = style[`border${side}Color`];
  const width = parseFloat(widthStr) || 0;
  if (width === 0 || styleStr === 'none' || styleStr === 'hidden') {
    return null;
  }

  const color = parseColor(colorStr);
  if (!color.hex || color.opacity === 0) return null;

  let dash = 'solid';
  if (styleStr === 'dashed') dash = 'dash';
  if (styleStr === 'dotted') dash = 'dot';

  return {
    pt: width * 0.75 * scale,
    color: color.hex,
    type: dash,
  };
}

/**
 * Extracts native table data for PptxGenJS.
 */
export function extractTableData(node, scale) {
  const rows = [];
  const colWidths = [];

  const firstRow = node.querySelector('tr');
  if (firstRow) {
    const cells = Array.from(firstRow.children);
    cells.forEach((cell) => {
      const rect = cell.getBoundingClientRect();
      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const wIn = (rect.width * (1 / 96) * scale) / colspan;
      for (let i = 0; i < colspan; i++) {
        colWidths.push(wIn);
      }
    });
  }

  const tableStyle = window.getComputedStyle(node);
  const borderSpacing = tableStyle.borderSpacing.split(' ');
  const hSpace = parseFloat(borderSpacing[0]) || 0;
  const vSpace = parseFloat(borderSpacing[1] || borderSpacing[0]) || 0;
  const hSpacePt = hSpace * 0.75 * scale;
  const vSpacePt = vSpace * 0.75 * scale;

  const trList = node.querySelectorAll('tr');
  trList.forEach((tr) => {
    const rowData = [];
    const cellList = Array.from(tr.children).filter((c) => ['TD', 'TH'].includes(c.tagName));
    cellList.forEach((cell) => {
      const style = window.getComputedStyle(cell);
      const cellParts = collectTextParts(cell, style, scale);
      const cellText =
        cellParts && cellParts.length > 0
          ? cellParts
          : cell.innerText.replace(/[\n\r\t]+/g, ' ').trim();

      const textStyle = getTextStyle(style, scale);

      let bg = parseColor(style.backgroundColor);
      if ((!bg.hex || bg.opacity === 0) && style.backgroundImage && style.backgroundImage !== 'none') {
        const fallback = getGradientFallbackColor(style.backgroundImage);
        if (fallback) bg = parseColor(fallback);
      }
      const fill = bg.hex && bg.opacity > 0 ? { color: bg.hex } : null;

      let align = 'left';
      if (style.textAlign === 'center') align = 'center';
      if (style.textAlign === 'right' || style.textAlign === 'end') align = 'right';

      let valign = 'top';
      if (style.verticalAlign === 'middle') valign = 'middle';
      if (style.verticalAlign === 'bottom') valign = 'bottom';

      const padding = getPadding(style, scale);
      const margin = [
        padding[0] * 72 + vSpacePt / 2,
        padding[1] * 72 + hSpacePt / 2,
        padding[2] * 72 + vSpacePt / 2,
        padding[3] * 72 + hSpacePt / 2,
      ];

      const borderTop = getTableBorder(style, 'Top', scale);
      const borderRight = getTableBorder(style, 'Right', scale);
      const borderBottom = getTableBorder(style, 'Bottom', scale);
      const borderLeft = getTableBorder(style, 'Left', scale);

      rowData.push({
        text: cellText,
        options: {
          color: textStyle.color,
          fontFace: textStyle.fontFace,
          fontSize: textStyle.fontSize,
          bold: textStyle.bold,
          italic: textStyle.italic,
          underline: textStyle.underline,
          fill,
          align,
          valign,
          margin,
          rowspan: parseInt(cell.getAttribute('rowspan')) || null,
          colspan: parseInt(cell.getAttribute('colspan')) || null,
          border: [borderTop, borderRight, borderBottom, borderLeft],
        },
      });
    });

    if (rowData.length > 0) rows.push(rowData);
  });

  return { rows, colWidths };
}

export function isClippedByParent(node) {
  let parent = node.parentElement;
  while (parent && parent !== document.body) {
    const style = window.getComputedStyle(parent);
    const overflow = style.overflow;
    if (overflow === 'hidden' || overflow === 'clip') return true;
    parent = parent.parentElement;
  }
  return false;
}

export function getGradientFallbackColor(bgImage) {
  if (!bgImage || bgImage === 'none') return null;
  const match = bgImage.match(/gradient\((.*)\)/);
  if (!match) return null;
  const content = match[1];
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

  for (const part of parts) {
    if (/^(to\s|[\d.]+(deg|rad|turn|grad))/.test(part)) continue;
    const colorPart = part.replace(/\s+(-?[\d.]+(%|px|em|rem|ch|vh|vw)?)$/, '');
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
  const top = { width: parseFloat(style.borderTopWidth) || 0, style: style.borderTopStyle, color: parseColor(style.borderTopColor).hex };
  const right = { width: parseFloat(style.borderRightWidth) || 0, style: style.borderRightStyle, color: parseColor(style.borderRightColor).hex };
  const bottom = { width: parseFloat(style.borderBottomWidth) || 0, style: style.borderBottomStyle, color: parseColor(style.borderBottomColor).hex };
  const left = { width: parseFloat(style.borderLeftWidth) || 0, style: style.borderLeftStyle, color: parseColor(style.borderLeftColor).hex };

  const hasAnyBorder = top.width > 0 || right.width > 0 || bottom.width > 0 || left.width > 0;
  if (!hasAnyBorder) return { type: 'none' };

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
        transparency: (1 - parseColor(style.borderTopColor).opacity) * 100,
        dashType: mapDashType(top.style),
      },
    };
  }

  return { type: 'composite', sides: { top, right, bottom, left } };
}

export function generateCompositeBorderSVG(w, h, radius, sides) {
  radius = radius / 2;
  const clipId = 'clip_' + Math.random().toString(36).substr(2, 9);
  let borderRects = '';
  if (sides.top.width > 0 && sides.top.color) {
    borderRects += `<rect x="0" y="0" width="${w}" height="${sides.top.width}" fill="#${sides.top.color}" />`;
  }
  if (sides.right.width > 0 && sides.right.color) {
    borderRects += `<rect x="${w - sides.right.width}" y="0" width="${sides.right.width}" height="${h}" fill="#${sides.right.color}" />`;
  }
  if (sides.bottom.width > 0 && sides.bottom.color) {
    borderRects += `<rect x="0" y="${h - sides.bottom.width}" width="${w}" height="${sides.bottom.width}" fill="#${sides.bottom.color}" />`;
  }
  if (sides.left.width > 0 && sides.left.color) {
    borderRects += `<rect x="0" y="0" width="${sides.left.width}" height="${h}" fill="#${sides.left.color}" />`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><defs><clipPath id="${clipId}"><rect x="0" y="0" width="${w}" height="${h}" rx="${radius}" ry="${radius}" /></clipPath></defs><g clip-path="url(#${clipId})">${borderRects}</g></svg>`;
  return 'data:image/svg+xml;base64,' + btoa(svg);
}

export function generateCustomShapeSVG(w, h, color, opacity, radii) {
  let { tl, tr, br, bl } = radii;
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
  const path = `M ${tl} 0 L ${w - tr} 0 A ${tr} ${tr} 0 0 1 ${w} ${tr} L ${w} ${h - br} A ${br} ${br} 0 0 1 ${w - br} ${h} L ${bl} ${h} A ${bl} ${bl} 0 0 1 0 ${h - bl} L 0 ${tl} A ${tl} ${tl} 0 0 1 ${tl} 0 Z`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><path d="${path}" fill="#${color}" fill-opacity="${opacity}" /></svg>`;
  return 'data:image/svg+xml;base64,' + btoa(svg);
}

export function parseColor(str) {
  if (!str || str === 'transparent' || str.trim?.() === 'rgba(0, 0, 0, 0)') {
    return { hex: null, opacity: 0 };
  }

  const ctx = getCtx();
  ctx.fillStyle = str;
  const computed = ctx.fillStyle;

  if (computed.startsWith('#')) {
    let hex = computed.slice(1);
    let opacity = 1;
    if (hex.length === 3 || hex.length === 4) {
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    }
    if (hex.length === 8) {
      opacity = parseInt(hex.slice(6), 16) / 255;
      hex = hex.slice(0, 6);
    }
    return { hex: hex.toUpperCase(), opacity };
  }

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

  ctx.clearRect(0, 0, 1, 1);
  ctx.fillRect(0, 0, 1, 1);
  const data = ctx.getImageData(0, 0, 1, 1).data;
  const r = data[0];
  const g = data[1];
  const b = data[2];
  const a = data[3] / 255;
  if (a === 0) return { hex: null, opacity: 0 };
  const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  return { hex, opacity: a };
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
  let colorObj = parseColor(style.color);
  let opacity = colorObj.opacity * inheritedOpacity;
  const elOpacity = parseFloat(style.opacity);
  if (!isNaN(elOpacity)) opacity *= elOpacity;

  const bgClip = style.webkitBackgroundClip || style.backgroundClip;
  if (colorObj.opacity === 0 && bgClip === 'text') {
    const fallback = getGradientFallbackColor(style.backgroundImage);
    if (fallback) colorObj = parseColor(fallback);
  }

  let lineSpacing = null;
  const fontSizePx = parseFloat(style.fontSize);
  const lhStr = style.lineHeight;
  if (lhStr && lhStr !== 'normal') {
    let lhPx = parseFloat(lhStr);
    if (/^[0-9.]+$/.test(lhStr)) lhPx *= fontSizePx;
    if (!isNaN(lhPx) && lhPx > 0) lineSpacing = lhPx * 0.75 * scale;
  }

  let paraSpaceBefore = 0;
  let paraSpaceAfter = 0;
  if (includeMargins) {
    const mt = parseFloat(style.marginTop) || 0;
    const mb = parseFloat(style.marginBottom) || 0;
    if (mt > 0) paraSpaceBefore = mt * 0.75 * scale;
    if (mb > 0) paraSpaceAfter = mb * 0.75 * scale;
  }

  const transparency = Math.round((1 - opacity) * 100);
  return {
    color: colorObj.hex || '000000',
    ...(transparency > 0 && { transparency }),
    fontFace: style.fontFamily.split(',')[0].replace(/['"]/g, ''),
    fontSize: Math.floor(fontSizePx * 0.75 * scale * 10) / 10,
    bold: parseInt(style.fontWeight) >= 600,
    italic: style.fontStyle === 'italic',
    underline: style.textDecoration.includes('underline'),
    ...(lineSpacing && { lineSpacing }),
    ...(paraSpaceBefore > 0 && { paraSpaceBefore }),
    ...(paraSpaceAfter > 0 && { paraSpaceAfter }),
    ...(parseColor(style.backgroundColor).hex ? { highlight: parseColor(style.backgroundColor).hex } : {}),
    ...(style.letterSpacing && style.letterSpacing !== 'normal'
      ? { charSpacing: parseFloat(style.letterSpacing) * 0.75 * scale }
      : {}),
  };
}

export function isTextContainer(node) {
  const hasText = node.textContent.trim().length > 0;
  if (!hasText) return false;
  const children = Array.from(node.children);
  if (children.length === 0) return true;

  const isSafeInline = (el) => {
    if (el.tagName.includes('-')) return false;
    if (el.tagName === 'IMG' || el.tagName === 'SVG') return false;
    if (el.tagName === 'I' || el.tagName === 'SPAN') {
      const cls = el.getAttribute('class') || '';
      if (
        cls.includes('fa-') ||
        cls.includes('fas') ||
        cls.includes('far') ||
        cls.includes('fab') ||
        cls.includes('material-icons') ||
        cls.includes('bi-') ||
        cls.includes('icon')
      ) {
        return false;
      }
    }

    const style = window.getComputedStyle(el);
    const display = style.display;
    const isInlineTag = ['SPAN', 'B', 'STRONG', 'EM', 'I', 'A', 'SMALL', 'MARK'].includes(el.tagName);
    const isInlineDisplay = display.includes('inline');
    if (!isInlineTag && !isInlineDisplay) return false;

    const bgColor = parseColor(style.backgroundColor);
    const hasVisibleBg = bgColor.hex && bgColor.opacity > 0;
    const hasBorder = parseFloat(style.borderWidth) > 0 && parseColor(style.borderColor).opacity > 0;
    const hasContent = el.textContent.trim().length > 0;
    if (!hasContent && (hasVisibleBg || hasBorder)) return false;
    return true;
  };

  return children.every(isSafeInline);
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

export function svgToSvg(node) {
  return new Promise((resolve) => {
    try {
      const clone = node.cloneNode(true);
      const rect = node.getBoundingClientRect();
      const width = rect.width || 300;
      const height = rect.height || 150;
      inlineSvgStyles(node, clone);
      clone.setAttribute('width', width);
      clone.setAttribute('height', height);
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      if (clone.querySelector('[*|href]') || clone.innerHTML.includes('xlink:')) {
        clone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
      }
      const xml = new XMLSerializer().serializeToString(clone);
      const svgUrl = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(xml)))}`;
      resolve(svgUrl);
    } catch (e) {
      console.warn('SVG serialization failed:', e);
      resolve(null);
    }
  });
}

function inlineSvgStyles(source, target) {
  const computed = window.getComputedStyle(source);
  const properties = [
    'fill',
    'stroke',
    'stroke-width',
    'stroke-linecap',
    'stroke-linejoin',
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
      const val = computed[prop];
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
        angle,
        blur: blur * 0.75 * scale,
        offset: distance * 0.75 * scale,
        color: colorObj.hex || '000000',
        opacity: colorObj.opacity,
      };
    }
  }
  return null;
}

export function generateGradientSVG(w, h, bgString, radius, border) {
  try {
    const match = bgString.match(/linear-gradient\((.*)\)/);
    if (!match) return null;
    const content = match[1];
    const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
    if (parts.length < 2) return null;

    let x1 = '0%', y1 = '0%', x2 = '0%', y2 = '100%';
    let stopsStartIndex = 0;
    const firstPart = parts[0].toLowerCase();

    if (firstPart.startsWith('to ')) {
      stopsStartIndex = 1;
      const direction = firstPart.replace('to ', '').trim();
      switch (direction) {
        case 'top':
          y1 = '100%';
          y2 = '0%';
          break;
        case 'bottom':
          y1 = '0%';
          y2 = '100%';
          break;
        case 'left':
          x1 = '100%';
          x2 = '0%';
          break;
        case 'right':
          x2 = '100%';
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
    } else if (firstPart.match(/^-?[\d.]+(deg|rad|turn|grad)$/)) {
      stopsStartIndex = 1;
      const val = parseFloat(firstPart);
      if (!isNaN(val)) {
        const deg = firstPart.includes('rad') ? val * (180 / Math.PI) : val;
        const cssRad = ((deg - 90) * Math.PI) / 180;
        const gradientScale = 50;
        const cos = Math.cos(cssRad);
        const sin = Math.sin(cssRad);
        x1 = (50 - sin * gradientScale).toFixed(1) + '%';
        y1 = (50 + cos * gradientScale).toFixed(1) + '%';
        x2 = (50 + sin * gradientScale).toFixed(1) + '%';
        y2 = (50 - cos * gradientScale).toFixed(1) + '%';
      }
    }

    let stopsXML = '';
    const stopParts = parts.slice(stopsStartIndex);
    stopParts.forEach((part, idx) => {
      let color = part;
      let offset = Math.round((idx / (stopParts.length - 1)) * 100) + '%';
      const posMatch = part.match(/^(.*?)\s+(-?[\d.]+(?:%|px)?)$/);
      if (posMatch) {
        color = posMatch[1];
        offset = posMatch[2];
      }
      let opacity = 1;
      if (color.includes('rgba')) {
        const rgbaMatch = color.match(/[\d.]+/g);
        if (rgbaMatch && rgbaMatch.length >= 4) {
          opacity = rgbaMatch[3];
          color = `rgb(${rgbaMatch[0]},${rgbaMatch[1]},${rgbaMatch[2]})`;
        }
      }
      stopsXML += `<stop offset="${offset}" stop-color="${color}" stop-opacity="${opacity}" />`;
    });

    let strokeAttr = '';
    if (border) strokeAttr = `stroke="#${border.color}" stroke-width="${border.width}"`;

    let tl = 0, tr = 0, br = 0, bl = 0;
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

    const pathD = `M ${tl} 0 L ${w - tr} 0 A ${tr} ${tr} 0 0 1 ${w} ${tr} L ${w} ${h - br} A ${br} ${br} 0 0 1 ${w - br} ${h} L ${bl} ${h} A ${bl} ${bl} 0 0 1 0 ${h - bl} L 0 ${tl} A ${tl} ${tl} 0 0 1 ${tl} 0 Z`;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><defs><linearGradient id="g" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stopsXML}</linearGradient></defs><path d="${pathD}" fill="url(#g)" ${strokeAttr}/></svg>`;
    return 'data:image/svg+xml;base64,' + btoa(svg);
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
  let shapeTag = '';
  const isCircle = radius >= Math.min(w, h) / 2 - 1 && Math.abs(w - h) < 2;
  if (isCircle) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    shapeTag = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="#${color}" filter="url(#blur)" />`;
  } else {
    shapeTag = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${radius}" ry="${radius}" fill="#${color}" filter="url(#blur)" />`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${fullW}" height="${fullH}" viewBox="0 0 ${fullW} ${fullH}"><defs><filter id="blur"><feGaussianBlur stdDeviation="${blurPx}" /></filter></defs>${shapeTag}</svg>`;
  return { data: 'data:image/svg+xml;base64,' + btoa(svg), padding };
}

export function getUsedFontFamilies(root) {
  const families = new Set();
  function scan(node) {
    if (node.nodeType === 1) {
      const style = window.getComputedStyle(node);
      const fontList = style.fontFamily.split(',');
      const primary = fontList[0].trim().replace(/['"]/g, '');
      if (primary) families.add(primary);
    }
    for (const child of node.childNodes) scan(child);
  }

  const elements = Array.isArray(root) ? root : [root];
  elements.forEach((el) => {
    const node = typeof el === 'string' ? document.querySelector(el) : el;
    if (node) scan(node);
  });
  return families;
}

export async function getAutoDetectedFonts(usedFamilies) {
  const foundFonts = [];
  const processedUrls = new Set();
  const extractUrl = (srcStr) => {
    const matches = srcStr.match(/url\((['"]?)(.*?)\1\)/g);
    if (!matches) return null;
    let chosenUrl = null;
    for (const match of matches) {
      const urlRaw = match.replace(/url\((['"]?)(.*?)\1\)/, '$2');
      if (urlRaw.startsWith('data:')) continue;
      if (urlRaw.includes('.ttf') || urlRaw.includes('.otf') || urlRaw.includes('.woff')) {
        chosenUrl = urlRaw;
        break;
      }
      if (!chosenUrl) chosenUrl = urlRaw;
    }
    return chosenUrl;
  };

  for (const sheet of Array.from(document.styleSheets)) {
    try {
      const rules = sheet.cssRules || sheet.rules;
      if (!rules) continue;
      for (const rule of Array.from(rules)) {
        if (rule.constructor.name === 'CSSFontFaceRule' || rule.type === 5) {
          const familyName = rule.style.getPropertyValue('font-family').replace(/['"]/g, '').trim();
          if (usedFamilies.has(familyName)) {
            const src = rule.style.getPropertyValue('src');
            const url = extractUrl(src);
            if (url && !processedUrls.has(url)) {
              processedUrls.add(url);
              foundFonts.push({ name: familyName, url });
            }
          }
        }
      }
    } catch (e) {
      console.warn('error:', e);
      console.warn('Cannot scan stylesheet for fonts (CORS restriction):', sheet.href);
    }
  }
  return foundFonts;
}

export function collectTextParts(node, parentStyle, scale, activeHyperlink = null, isRoot = true, inheritedOpacity = 1) {
  const parts = [];
  let hyperlink = activeHyperlink;
  if (!hyperlink && node.nodeType === 1) {
    const aNode = node.closest('a');
    if (aNode) {
      const href = aNode.getAttribute('href');
      if (href) hyperlink = { url: href, tooltip: aNode.getAttribute('title') || undefined };
    }
  }

  if (node.nodeType === 1) {
    const beforeStyle = window.getComputedStyle(node, '::before');
    const content = beforeStyle.content;
    if (content && content !== 'none' && content !== 'normal' && content !== '""') {
      const cleanContent = content.replace(/^["']|["']$/g, '');
      if (cleanContent.trim()) {
        const textOpts = getTextStyle(window.getComputedStyle(node), scale);
        if (hyperlink) textOpts.hyperlink = hyperlink;
        parts.push({ text: cleanContent + ' ', options: textOpts });
      }
    }
  }

  let trimNextLeading = false;
  node.childNodes.forEach((child, index) => {
    if (child.nodeType === 3) {
      let val = child.nodeValue.replace(/[\n\r\t]+/g, ' ').replace(/\s{2,}/g, ' ');
      if (index === 0) val = val.trimStart();
      if (trimNextLeading) {
        val = val.trimStart();
        trimNextLeading = false;
      }
      if (index === node.childNodes.length - 1) val = val.trimEnd();
      if (val) {
        const styleToUse = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
        const transform = styleToUse.textTransform;
        if (transform === 'uppercase') val = val.toUpperCase();
        else if (transform === 'lowercase') val = val.toLowerCase();
        else if (transform === 'capitalize') val = val.replace(/\b\w/g, (c) => c.toUpperCase());
        const textOpts = getTextStyle(styleToUse, scale, !isRoot, inheritedOpacity);
        if (hyperlink) textOpts.hyperlink = hyperlink;
        if (textOpts.highlight) delete textOpts.highlight;
        parts.push({ text: val, options: textOpts });
      }
    } else if (child.nodeType === 1) {
      if (child.tagName === 'BR') {
        if (parts.length > 0) {
          const lastPart = parts[parts.length - 1];
          if (lastPart.text && typeof lastPart.text === 'string') lastPart.text = lastPart.text.trimEnd();
        }
        parts.push({ text: '', options: { breakLine: true } });
        trimNextLeading = true;
      } else {
        const isBlock = ['DIV', 'P', 'LI'].includes(child.tagName);
        if (isBlock && parts.length > 0 && !parts[parts.length - 1].options?.breakLine) {
          parts.push({ text: '', options: { breakLine: true } });
        }
        const childParts = collectTextParts(child, parentStyle, scale, hyperlink, false, inheritedOpacity);
        if (childParts.length > 0) parts.push(...childParts);
        if (isBlock) {
          parts.push({ text: '', options: { breakLine: true } });
          trimNextLeading = true;
        }
      }
    }
  });

  while (parts.length > 0 && parts[parts.length - 1].options?.breakLine && parts[parts.length - 1].text === '') {
    parts.pop();
  }
  return parts;
}
