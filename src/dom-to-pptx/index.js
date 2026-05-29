import * as PptxGenJSImport from 'pptxgenjs';
import JSZip from 'jszip';

import { PPTXEmbedFonts } from './font-embedder.js';
import { getProcessedImage } from './image-processor.js';
import { normalizePptxZip } from './pptx-normalizer.js';
import {
  collectTextParts,
  extractTableData,
  getAutoDetectedFonts,
  getBorderInfo,
  getPadding,
  getRotation,
  getTextStyle,
  getUsedFontFamilies,
  getVisibleShadow,
  isTextContainer,
  parseColor,
  pxToInch,
} from './utils.js';

const PptxGenJS = PptxGenJSImport?.default ?? PptxGenJSImport;
const DEFAULT_WIDTH = 10;
const DEFAULT_HEIGHT = 5.625;

export async function exportToPptx(target, options = {}) {
  const PptxConstructor = resolvePptxConstructor(PptxGenJS);
  if (!PptxConstructor) throw new Error('PptxGenJS constructor not found.');

  const pptx = new PptxConstructor();
  const layout = applyLayout(pptx, options);
  const elements = Array.isArray(target) ? target : [target];

  for (const entry of elements) {
    const root = typeof entry === 'string' ? document.querySelector(entry) : entry;
    if (!root) continue;
    const slide = pptx.addSlide();
    await processSlide(root, slide, { ...options, ...layout });
  }

  let blob = await pptx.write({ outputType: 'blob' });
  const fontsToEmbed = await resolveFonts(elements, options);

  if (fontsToEmbed.length || options.skipNormalize !== true) {
    const zip = await JSZip.loadAsync(blob);
    if (fontsToEmbed.length) {
      const embedder = new PPTXEmbedFonts();
      await embedder.loadZip(zip);
      for (const font of fontsToEmbed) {
        const response = await fetch(font.url);
        const buffer = await response.arrayBuffer();
        const type = inferFontType(font.url);
        await embedder.addFont(font.name, buffer, type);
      }
      await embedder.updateFiles();
      if (options.skipNormalize !== true) await normalizePptxZip(zip);
      blob = await embedder.generateBlob();
    } else {
      await normalizePptxZip(zip);
      blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    }
  }

  if (!options.skipDownload) downloadBlob(blob, options.fileName || 'export.pptx');
  return blob;
}

async function processSlide(root, slide, options) {
  const rootRect = root.getBoundingClientRect();
  const contentWidthIn = pxToInch(rootRect.width);
  const contentHeightIn = pxToInch(rootRect.height);
  const scale = Math.min(options._slideWidth / contentWidthIn, options._slideHeight / contentHeightIn);
  const context = {
    root,
    rootRect,
    scale,
    offX: (options._slideWidth - contentWidthIn * scale) / 2,
    offY: (options._slideHeight - contentHeightIn * scale) / 2,
  };

  await renderElement(root, slide, context, options, 1);
}

async function renderElement(node, slide, context, options, inheritedOpacity) {
  if (node.nodeType !== 1) return;
  const style = window.getComputedStyle(node);
  if (style.display === 'none' || style.visibility === 'hidden') return;

  const opacity = inheritedOpacity * (parseFloat(style.opacity) || 1);
  const rect = node.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) return;

  const box = toPptxBox(rect, context, style);
  const tag = node.tagName;

  if (tag === 'IMG' || tag === 'SVG' || tag === 'CANVAS') {
    const data = await getProcessedImage(node, options);
    if (data) slide.addImage({ data, ...box, transparency: Math.round((1 - opacity) * 100) });
    return;
  }

  if (tag === 'TABLE') {
    const tableData = extractTableData(node, context.scale);
    slide.addTable(tableData.rows, { ...box, border: { type: 'none' } });
    return;
  }

  addBackgroundShape(slide, style, box, opacity);

  if (isTextContainer(node)) {
    const padding = getPadding(style);
    const textBox = {
      ...box,
      x: box.x + pxToInch(padding.left, context.scale),
      y: box.y + pxToInch(padding.top, context.scale),
      w: Math.max(0.01, box.w - pxToInch(padding.left + padding.right, context.scale)),
      h: Math.max(0.01, box.h - pxToInch(padding.top + padding.bottom, context.scale)),
      margin: 0,
      rotate: box.rotate,
      breakLine: false,
    };
    const textParts = collectTextParts(node, context.scale, opacity);
    slide.addText(textParts.length ? textParts : node.textContent.trim(), {
      ...textBox,
      ...getTextStyle(style, context.scale, opacity),
      fit: 'shrink',
    });
    return;
  }

  for (const child of node.childNodes) {
    await renderElement(child, slide, context, options, opacity);
  }
}

function addBackgroundShape(slide, style, box, opacity) {
  const fill = parseColor(style.backgroundColor, 'FFFFFF');
  const hasFill = fill.transparency < 100;
  const border = getBorderInfo(style);

  if (!hasFill && !border) return;

  slide.addShape('rect', {
    ...box,
    fill: hasFill
      ? { color: fill.color, transparency: Math.min(100, fill.transparency + Math.round((1 - opacity) * 100)) }
      : { color: 'FFFFFF', transparency: 100 },
    line: border || { transparency: 100 },
    shadow: getVisibleShadow(style),
  });
}

function toPptxBox(rect, context, style) {
  return {
    x: context.offX + pxToInch(rect.left - context.rootRect.left, context.scale),
    y: context.offY + pxToInch(rect.top - context.rootRect.top, context.scale),
    w: pxToInch(rect.width, context.scale),
    h: pxToInch(rect.height, context.scale),
    rotate: getRotation(style.transform),
  };
}

function applyLayout(pptx, options) {
  if (options.width && options.height) {
    pptx.defineLayout({ name: 'CUSTOM', width: options.width, height: options.height });
    pptx.layout = 'CUSTOM';
    return { _slideWidth: options.width, _slideHeight: options.height };
  }

  pptx.layout = options.layout || 'LAYOUT_16x9';
  if (options.layout === 'LAYOUT_4x3') return { _slideWidth: 10, _slideHeight: 7.5 };
  if (options.layout === 'LAYOUT_16x10') return { _slideWidth: 10, _slideHeight: 6.25 };
  if (options.layout === 'LAYOUT_WIDE') return { _slideWidth: 13.333, _slideHeight: 7.5 };
  return { _slideWidth: DEFAULT_WIDTH, _slideHeight: DEFAULT_HEIGHT };
}

async function resolveFonts(elements, options) {
  const fonts = [...(options.fonts || [])];
  if (options.autoEmbedFonts) {
    const usedFamilies = getUsedFontFamilies(elements);
    const detectedFonts = await getAutoDetectedFonts(usedFamilies);
    const knownNames = new Set(fonts.map((font) => font.name));
    detectedFonts.forEach((font) => {
      if (!knownNames.has(font.name)) fonts.push(font);
    });
  }
  return fonts;
}

function inferFontType(url) {
  const ext = url.split('.').pop().split(/[?#]/)[0].toLowerCase();
  if (['woff', 'otf'].includes(ext)) return ext;
  return 'ttf';
}

function resolvePptxConstructor(pkg) {
  if (typeof pkg === 'function') return pkg;
  if (typeof pkg?.default === 'function') return pkg.default;
  if (typeof pkg?.PptxGenJS === 'function') return pkg.PptxGenJS;
  if (typeof pkg?.PptxGenJS?.default === 'function') return pkg.PptxGenJS.default;
  return null;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}
