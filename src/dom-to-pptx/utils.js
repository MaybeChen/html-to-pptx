const PPI = 96;

export const pxToInch = (px, scale = 1) => (px / PPI) * scale;

export function parseColor(value, fallback = 'FFFFFF') {
  if (!value || value === 'transparent' || value === 'rgba(0, 0, 0, 0)') {
    return { color: fallback, transparency: 100 };
  }

  const hex = value.trim().match(/^#([0-9a-f]{3,8})$/i);
  if (hex) {
    let normalized = hex[1];
    if (normalized.length === 3 || normalized.length === 4) {
      normalized = normalized
        .split('')
        .map((char) => char + char)
        .join('');
    }

    const alpha = normalized.length === 8 ? parseInt(normalized.slice(6, 8), 16) / 255 : 1;
    return {
      color: normalized.slice(0, 6).toUpperCase(),
      transparency: Math.round((1 - alpha) * 100),
    };
  }

  const rgb = value.match(/^rgba?\(([^)]+)\)$/i);
  if (rgb) {
    const [red, green, blue, alpha = '1'] = rgb[1].split(',').map((part) => part.trim());
    const toHex = (channel) => Number(channel).toString(16).padStart(2, '0');
    return {
      color: `${toHex(red)}${toHex(green)}${toHex(blue)}`.toUpperCase(),
      transparency: Math.round((1 - Number(alpha)) * 100),
    };
  }

  return { color: fallback, transparency: 0 };
}

export function getRotation(transform) {
  if (!transform || transform === 'none') return 0;
  const matrix = transform.match(/matrix\(([^)]+)\)/);
  if (!matrix) return 0;
  const [a, b] = matrix[1].split(',').map(Number);
  return Math.round(Math.atan2(b, a) * (180 / Math.PI));
}

export function getPadding(style) {
  return {
    top: parseFloat(style.paddingTop) || 0,
    right: parseFloat(style.paddingRight) || 0,
    bottom: parseFloat(style.paddingBottom) || 0,
    left: parseFloat(style.paddingLeft) || 0,
  };
}

export function getBorderInfo(style) {
  const width = parseFloat(style.borderTopWidth) || 0;
  if (width <= 0 || style.borderTopStyle === 'none') return undefined;
  const parsed = parseColor(style.borderTopColor, '000000');
  return { color: parsed.color, width: Math.max(width * 0.75, 0.25), transparency: parsed.transparency };
}

export function getVisibleShadow(style) {
  if (!style.boxShadow || style.boxShadow === 'none') return undefined;
  const color = style.boxShadow.match(/rgba?\([^)]+\)|#[0-9a-f]{3,8}/i)?.[0];
  const numbers = style.boxShadow.replace(/rgba?\([^)]+\)|#[0-9a-f]{3,8}/gi, '').match(/-?\d+(\.\d+)?px/g);
  if (!numbers || numbers.length < 2) return undefined;
  const [offsetX, offsetY, blur = '0px'] = numbers.map((part) => parseFloat(part));
  const distance = Math.sqrt(offsetX ** 2 + offsetY ** 2) * 0.75;
  const angle = Math.round((Math.atan2(offsetY, offsetX) * 180) / Math.PI + 90);
  const parsed = parseColor(color || 'rgba(0,0,0,0.3)', '000000');
  return { type: 'outer', color: parsed.color, opacity: 1 - parsed.transparency / 100, blur, angle, distance };
}

export function getTextStyle(style, scale = 1, inheritedOpacity = 1) {
  const color = parseColor(style.color, '000000');
  const fontSizePx = parseFloat(style.fontSize) || 16;
  const isBold = Number(style.fontWeight) >= 600 || style.fontWeight === 'bold';

  return {
    fontFace: normalizeFontFamily(style.fontFamily),
    fontSize: Math.max(1, fontSizePx * 0.75 * scale),
    color: color.color,
    transparency: Math.min(100, Math.round(color.transparency + (1 - inheritedOpacity) * 100)),
    bold: isBold,
    italic: style.fontStyle === 'italic',
    underline: style.textDecorationLine?.includes('underline') || false,
    breakLine: false,
  };
}

export function normalizeFontFamily(fontFamily = '') {
  const first = fontFamily.split(',')[0]?.replaceAll('"', '').replaceAll("'", '').trim();
  if (!first || ['system-ui', 'ui-sans-serif', '-apple-system', 'BlinkMacSystemFont'].includes(first)) {
    return 'Arial';
  }
  return first;
}

export function isTextContainer(node) {
  if (!node || node.nodeType !== 1) return false;
  const text = Array.from(node.childNodes).some(
    (child) => child.nodeType === 3 && child.nodeValue.trim().length > 0
  );
  return text && !['IMG', 'SVG', 'CANVAS', 'TABLE'].includes(node.tagName);
}

export function collectTextParts(node, scale = 1, inheritedOpacity = 1) {
  const parts = [];
  const visit = (current) => {
    if (current.nodeType === 3) {
      const text = current.nodeValue.replace(/\s+/g, ' ');
      if (text.trim()) {
        parts.push({ text, options: getTextStyle(window.getComputedStyle(current.parentElement), scale, inheritedOpacity) });
      }
      return;
    }
    if (current.nodeType !== 1) return;
    for (const child of current.childNodes) visit(child);
  };
  visit(node);
  return parts.length ? parts : [{ text: node.textContent.trim(), options: getTextStyle(window.getComputedStyle(node), scale, inheritedOpacity) }];
}

export function getUsedFontFamilies(elements) {
  const families = new Set();
  const visit = (node) => {
    if (!node || node.nodeType !== 1) return;
    families.add(normalizeFontFamily(window.getComputedStyle(node).fontFamily));
    for (const child of node.children) visit(child);
  };
  elements.forEach((entry) => visit(typeof entry === 'string' ? document.querySelector(entry) : entry));
  return Array.from(families).filter(Boolean);
}

export async function getAutoDetectedFonts() {
  return [];
}

export function extractTableData(table, scale = 1) {
  const rows = Array.from(table.rows).map((row) =>
    Array.from(row.cells).map((cell) => ({
      text: cell.textContent.trim(),
      options: getTextStyle(window.getComputedStyle(cell), scale),
    }))
  );
  return { rows, colWidths: [] };
}

export function getWritingModeVert(writingMode) {
  return writingMode?.startsWith('vertical') || false;
}

export function getSoftEdges() {
  return undefined;
}

export function isClippedByParent() {
  return false;
}

export function generateGradientSVG() {
  return undefined;
}

export function generateBlurredSVG() {
  return undefined;
}

export function generateCompositeBorderSVG() {
  return undefined;
}

export function generateCustomShapeSVG() {
  return undefined;
}

export function svgToSvg(node) {
  return `data:image/svg+xml;base64,${btoa(new XMLSerializer().serializeToString(node))}`;
}

export async function svgToPng(node) {
  return svgToSvg(node);
}
