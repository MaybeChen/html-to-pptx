// src/index.js
import * as PptxGenJSImport from 'pptxgenjs';
import html2canvas from 'html2canvas';

// Normalize import
const PptxGenJS = PptxGenJSImport?.default ?? PptxGenJSImport;

import {
  parseColor,
  getTextStyle,
  isTextContainer,
  getVisibleShadow,
  generateGradientSVG,
  getRotation,
  getWritingModeVert,
  svgToPng,
  svgToSvg,
  getPadding,
  getSoftEdges,
  generateBlurredSVG,
  getBorderInfo,
  generateCompositeBorderSVG,
  isClippedByParent,
  generateCustomShapeSVG,
  getUsedFontFamilies,
  getAutoDetectedFonts,
  extractTableData,
  collectTextParts,
} from './utils.js';
import { getProcessedImage } from './image-processor.js';

const PPI = 96;
const PX_TO_INCH = 1 / PPI;

/**
 * Main export function.
 * @param {HTMLElement | string | Array} target
 * @param {Object} options
 * @param {string} [options.fileName]
 * @param {boolean} [options.skipDownload=false] - If true, prevents automatic download
 * @param {Object} [options.listConfig] - Config for bullets
 * @param {boolean} [options.svgAsVector=false] - If true, keeps SVG as vector (for Convert to Shape in PowerPoint)
 * @param {boolean} [options.skipNormalize=false] - If true, skips re-zipping with DEFLATE
 * and stripping dangling [Content_Types].xml Overrides. Leave it false unless you are
 * debugging the raw PptxGenJS output, otherwise Microsoft PowerPoint may reject the file.
 * @returns {Promise<Blob>} - Returns the generated PPTX Blob
 */
export async function exportToPptx(target, options = {}) {
  const resolvePptxConstructor = (pkg) => {
    if (!pkg) return null;
    if (typeof pkg === 'function') return pkg;
    if (pkg && typeof pkg.default === 'function') return pkg.default;
    if (pkg && typeof pkg.PptxGenJS === 'function') return pkg.PptxGenJS;
    if (pkg && pkg.PptxGenJS && typeof pkg.PptxGenJS.default === 'function')
      return pkg.PptxGenJS.default;
    return null;
  };

  const PptxConstructor = resolvePptxConstructor(PptxGenJS);
  if (!PptxConstructor) throw new Error('PptxGenJS constructor not found.');
  const pptx = new PptxConstructor();

  // 1. Layout Handling
  let finalWidth = 10; // default 16:9
  let finalHeight = 5.625;

  if (options.width && options.height) {
    pptx.defineLayout({ name: 'CUSTOM', width: options.width, height: options.height });
    pptx.layout = 'CUSTOM';
    finalWidth = options.width;
    finalHeight = options.height;
  } else if (options.layout) {
    pptx.layout = options.layout;
    // Map standard layouts for internal scale calculation if possible,
    // though PptxGenJS defaults to 16:9 if unknown.
    if (options.layout === 'LAYOUT_4x3') {
      finalWidth = 10;
      finalHeight = 7.5;
    } else if (options.layout === 'LAYOUT_16x10') {
      finalWidth = 10;
      finalHeight = 6.25;
    } else if (options.layout === 'LAYOUT_WIDE') {
      finalWidth = 13.3;
      finalHeight = 7.5;
    }
  } else {
    pptx.layout = 'LAYOUT_16x9';
  }

  // Pass these dimensions to options so processSlide can use them
  const extendedOptions = {
    ...options,
    _slideWidth: finalWidth,
    _slideHeight: finalHeight,
  };

  const elements = Array.isArray(target) ? target : [target];

  for (const el of elements) {
    const root = typeof el === 'string' ? document.querySelector(el) : el;
    if (!root) {
      console.warn('Element not found, skipping slide:', el);
      continue;
    }
    const slide = pptx.addSlide();
    await processSlide(root, slide, pptx, extendedOptions);
  }

  // 3. Font Embedding Logic
  let finalBlob;
  let fontsToEmbed = options.fonts || [];

  if (options.autoEmbedFonts) {
    // A. Scan DOM for used font families
    const usedFamilies = getUsedFontFamilies(elements);

    // B. Scan CSS for URLs matches
    const detectedFonts = await getAutoDetectedFonts(usedFamilies);

    // C. Merge (Avoid duplicates)
    const explicitNames = new Set(fontsToEmbed.map((f) => f.name));
    for (const autoFont of detectedFonts) {
      if (!explicitNames.has(autoFont.name)) {
        fontsToEmbed.push(autoFont);
      }
    }

    if (detectedFonts.length > 0) {
      console.log(
        'Auto-detected fonts:',
        detectedFonts.map((f) => f.name)
      );
    }
  }

  if (fontsToEmbed.length > 0) {
    // Generate initial PPTX
    const initialBlob = await pptx.write({ outputType: 'blob' });

    // Load font embedding dependencies only when fonts are actually embedded.
    const [{ default: JSZip }, { PPTXEmbedFonts }, { normalizePptxZip }] = await Promise.all([
      import('jszip'),
      import('./font-embedder.js'),
      import('./pptx-normalizer.js'),
    ]);
    const zip = await JSZip.loadAsync(initialBlob);
    const embedder = new PPTXEmbedFonts();
    await embedder.loadZip(zip);

    // Fetch and Embed
    for (const fontCfg of fontsToEmbed) {
      try {
        const response = await fetch(fontCfg.url);
        if (!response.ok) throw new Error(`Failed to fetch ${fontCfg.url}`);
        const buffer = await response.arrayBuffer();

        // Infer type
        const ext = fontCfg.url.split('.').pop().split(/[?#]/)[0].toLowerCase();
        let type = 'ttf';
        if (['woff', 'otf'].includes(ext)) type = ext;

        await embedder.addFont(fontCfg.name, buffer, type);
      } catch (e) {
        console.warn(`Failed to embed font: ${fontCfg.name} (${fontCfg.url})`, e);
      }
    }

    await embedder.updateFiles();
    if (options.skipNormalize !== true) {
      await normalizePptxZip(zip);
    }
    finalBlob = await embedder.generateBlob();
  } else {
    // No fonts to embed — still re-zip with DEFLATE and strip dangling Overrides
    // so Microsoft PowerPoint accepts the file (PptxGenJS leaves both issues
    // unresolved on its own; see 错误诊断.md).
    const initialBlob = await pptx.write({ outputType: 'blob' });
    if (options.skipNormalize === true) {
      finalBlob = initialBlob;
    } else {
      const [{ default: JSZip }, { normalizePptxZip }] = await Promise.all([
        import('jszip'),
        import('./pptx-normalizer.js'),
      ]);
      const zip = await JSZip.loadAsync(initialBlob);
      await normalizePptxZip(zip);
      finalBlob = await zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      });
    }
  }

  // 4. Output Handling
  // If skipDownload is NOT true, proceed with browser download
  if (!options.skipDownload) {
    const fileName = options.fileName || 'export.pptx';
    const url = URL.createObjectURL(finalBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Always return the blob so the caller can use it (e.g. upload to server)
  return finalBlob;
}

/**
 * Worker function to process a single DOM element into a single PPTX slide.
 * @param {HTMLElement} root - The root element for this slide.
 * @param {PptxGenJS.Slide} slide - The PPTX slide object to add content to.
 * @param {PptxGenJS} pptx - The main PPTX instance.
 */
async function processSlide(root, slide, pptx, globalOptions = {}) {
  const rootRect = root.getBoundingClientRect();
  const PPTX_WIDTH_IN = globalOptions._slideWidth || 10;
  const PPTX_HEIGHT_IN = globalOptions._slideHeight || 5.625;

  const contentWidthIn = rootRect.width * PX_TO_INCH;
  const contentHeightIn = rootRect.height * PX_TO_INCH;
  const scale = Math.min(PPTX_WIDTH_IN / contentWidthIn, PPTX_HEIGHT_IN / contentHeightIn);

  const layoutConfig = {
    rootX: rootRect.x,
    rootY: rootRect.y,
    scale: scale,
    offX: (PPTX_WIDTH_IN - contentWidthIn * scale) / 2,
    offY: (PPTX_HEIGHT_IN - contentHeightIn * scale) / 2,
  };

  const renderQueue = [];
  const asyncTasks = []; // Queue for heavy operations (Images, Canvas)
  let domOrderCounter = 0;

  // Sync Traversal Function
  function collect(node, parentZIndex, parentOpacity = 1) {
    const order = domOrderCounter++;

    let currentZ = parentZIndex;
    let currentOpacity = parentOpacity;
    let nodeStyle = null;
    const nodeType = node.nodeType;

    if (nodeType === 1) {
      nodeStyle = window.getComputedStyle(node);
      const elOpacity = parseFloat(nodeStyle.opacity);
      if (!isNaN(elOpacity)) {
        currentOpacity *= elOpacity;
      }

      // Optimization: Skip completely hidden elements immediately
      if (
        nodeStyle.display === 'none' ||
        nodeStyle.visibility === 'hidden' ||
        currentOpacity === 0
      ) {
        return;
      }
      if (nodeStyle.zIndex !== 'auto') {
        currentZ = parseInt(nodeStyle.zIndex);
      }
    }

    // Prepare the item. If it needs async work, it returns a 'job'
    const result = prepareRenderItem(
      node,
      { ...layoutConfig, root },
      order,
      pptx,
      currentZ,
      nodeStyle,
      { ...globalOptions, _inheritedOpacity: parentOpacity }
    );

    if (result) {
      if (result.items) {
        // Push items immediately to queue (data might be missing but filled later)
        renderQueue.push(...result.items);
      }
      if (result.job) {
        // Push the promise-returning function to the task list
        asyncTasks.push(result.job);
      }
      if (result.stopRecursion) return;
    }

    // Recurse children synchronously
    const childNodes = node.childNodes;
    for (let i = 0; i < childNodes.length; i++) {
      collect(childNodes[i], currentZ, currentOpacity);
    }
  }

  // 1. Traverse and build the structure (Fast)
  collect(root, 0);

  // 2. Execute all heavy tasks in parallel (Fast)
  if (asyncTasks.length > 0) {
    await Promise.all(asyncTasks.map((task) => task()));
  }

  // 3. Cleanup and Sort
  // Remove items that failed to generate data (marked with skip)
  const finalQueue = renderQueue.filter(
    (item) => !item.skip && (item.type !== 'image' || item.options.data)
  );

  finalQueue.sort((a, b) => {
    if (a.zIndex !== b.zIndex) return a.zIndex - b.zIndex;
    return a.domOrder - b.domOrder;
  });

  // 4. Add to Slide
  for (const item of finalQueue) {
    if (item.type === 'shape') slide.addShape(item.shapeType, item.options);
    if (item.type === 'image') slide.addImage(item.options);
    if (item.type === 'text') slide.addText(item.textParts, item.options);
    if (item.type === 'table') {
      slide.addTable(item.tableData.rows, {
        x: item.options.x,
        y: item.options.y,
        w: item.options.w,
        colW: item.tableData.colWidths, // Essential for correct layout
        autoPage: false,
        // Remove default table styles so our extracted CSS applies cleanly
        border: { type: 'none' },
        fill: { color: 'FFFFFF', transparency: 100 },
      });
    }
  }
}

/**
 * Optimized html2canvas wrapper
 * Includes fix for cropped icons by adjusting styles in the cloned document.
 */
async function elementToCanvasImage(node, widthPx, heightPx) {
  return new Promise((resolve) => {
    // 1. Assign a temp ID to locate the node inside the cloned document
    const originalId = node.id;
    const tempId = 'pptx-capture-' + Math.random().toString(36).substr(2, 9);
    node.id = tempId;

    const width = Math.max(Math.ceil(widthPx), 1);
    const height = Math.max(Math.ceil(heightPx), 1);
    const style = window.getComputedStyle(node);

    // Add padding to the clone to capture spilling content (like extensive font glyphs)
    const padding = 10;

    html2canvas(node, {
      backgroundColor: null,
      logging: false,
      scale: 3, // Higher scale for sharper icons
      useCORS: true, // critical for external fonts/images
      width: width + padding * 2, // Capture a larger area
      height: height + padding * 2,
      x: -padding, // Offset capture to include the padding
      y: -padding,
      onclone: (clonedDoc) => {
        const clonedNode = clonedDoc.getElementById(tempId);
        if (clonedNode) {
          // --- FIX: CLIP & FONT ISSUES ---
          // Apply styles DIRECTLY to elements to ensure html2canvas picks them up
          // This avoids issues where icon font glyphs or transformed content gets cropped.
          clonedNode.style.overflow = 'visible';
          clonedNode.style.width = width + 'px';
          clonedNode.style.height = height + 'px';
          clonedNode.style.boxSizing = style.boxSizing;
          clonedNode.style.transform = 'none';

          const descendants = clonedNode.querySelectorAll('*');
          descendants.forEach((el) => {
            el.style.overflow = 'visible';
            el.style.textRendering = 'geometricPrecision';
            el.style.webkitFontSmoothing = 'antialiased';
          });
        }
      },
    })
      .then((canvas) => {
        if (originalId) node.id = originalId;
        else node.removeAttribute('id');
        resolve(canvas.toDataURL('image/png'));
      })
      .catch((err) => {
        if (originalId) node.id = originalId;
        else node.removeAttribute('id');
        console.warn('html2canvas failed:', err);
        resolve(null);
      });
  });
}

function prepareRenderItem(node, layout, domOrder, pptx, zIndex, style, options) {
  if (node.nodeType !== 1) return null;

  const rect = node.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) return null;

  const scale = layout.scale;
  const x = layout.offX + (rect.x - layout.rootX) * PX_TO_INCH * scale;
  const y = layout.offY + (rect.y - layout.rootY) * PX_TO_INCH * scale;
  const w = rect.width * PX_TO_INCH * scale;
  const h = rect.height * PX_TO_INCH * scale;
  const opacity = options._inheritedOpacity ?? 1;
  const rotate = getRotation(style.transform);
  const radiusPx = parseFloat(style.borderTopLeftRadius) || 0;
  const radius = radiusPx * scale;
  const base = { x, y, w, h, rotate };

  if (node.tagName === 'TABLE') {
    return {
      items: [
        {
          type: 'table',
          tableData: extractTableData(node, scale),
          options: base,
          zIndex,
          domOrder,
        },
      ],
      stopRecursion: true,
    };
  }

  if (node.tagName === 'IMG') {
    const item = { type: 'image', options: { ...base }, zIndex, domOrder };
    return {
      items: [item],
      job: async () => {
        const src = node.currentSrc || node.src;
        item.options.data = await getProcessedImage(
          src,
          rect.width,
          rect.height,
          radiusPx,
          style.objectFit,
          style.objectPosition
        );
        if (!item.options.data) item.skip = true;
      },
      stopRecursion: true,
    };
  }

  if (node.tagName === 'SVG') {
    const item = { type: 'image', options: { ...base }, zIndex, domOrder };
    return {
      items: [item],
      job: async () => {
        item.options.data = options.svgAsVector ? await svgToSvg(node) : await svgToPng(node);
        if (!item.options.data) item.skip = true;
      },
      stopRecursion: true,
    };
  }

  if (node.tagName === 'CANVAS') {
    const data = node.toDataURL('image/png');
    return {
      items: [{ type: 'image', options: { ...base, data }, zIndex, domOrder }],
      stopRecursion: true,
    };
  }

  const items = [];
  const bg = parseColor(style.backgroundColor);
  const borderInfo = getBorderInfo(style, scale);
  const shadow = getVisibleShadow(style.boxShadow, scale);
  const softEdge = getSoftEdges(style.filter, scale);
  const bgImage = style.backgroundImage;
  const hasGradient = bgImage && bgImage.includes('gradient(');
  const clipped = isClippedByParent(node);

  if (hasGradient) {
    const data = generateGradientSVG(rect.width, rect.height, bgImage, radiusPx, borderInfo.options);
    if (data) items.push({ type: 'image', options: { ...base, data }, zIndex, domOrder });
  } else if (bg.hex && bg.opacity > 0) {
    const fill = { color: bg.hex, transparency: Math.round((1 - bg.opacity * opacity) * 100) };
    const line = borderInfo.type === 'uniform' ? borderInfo.options : { transparency: 100 };
    if (radiusPx > 0 && (style.borderTopLeftRadius !== style.borderTopRightRadius || style.borderTopLeftRadius !== style.borderBottomRightRadius || style.borderTopLeftRadius !== style.borderBottomLeftRadius)) {
      const radii = {
        tl: parseFloat(style.borderTopLeftRadius) || 0,
        tr: parseFloat(style.borderTopRightRadius) || 0,
        br: parseFloat(style.borderBottomRightRadius) || 0,
        bl: parseFloat(style.borderBottomLeftRadius) || 0,
      };
      items.push({
        type: 'image',
        options: { ...base, data: generateCustomShapeSVG(rect.width, rect.height, bg.hex, bg.opacity, radii) },
        zIndex,
        domOrder,
      });
    } else {
      items.push({
        type: 'shape',
        shapeType: pptx.ShapeType.roundRect,
        options: { ...base, fill, line, radius, shadow, softEdge },
        zIndex,
        domOrder,
      });
    }
  } else if (borderInfo.type === 'uniform') {
    items.push({
      type: 'shape',
      shapeType: pptx.ShapeType.rect,
      options: { ...base, fill: { color: 'FFFFFF', transparency: 100 }, line: borderInfo.options },
      zIndex,
      domOrder,
    });
  }

  if (borderInfo.type === 'composite') {
    items.push({
      type: 'image',
      options: { ...base, data: generateCompositeBorderSVG(rect.width, rect.height, radiusPx, borderInfo.sides) },
      zIndex,
      domOrder: domOrder + 0.1,
    });
  }

  if (softEdge && bg.hex) {
    const blurred = generateBlurredSVG(rect.width, rect.height, bg.hex, radiusPx, softEdge);
    items.push({
      type: 'image',
      options: {
        x: x - blurred.padding * PX_TO_INCH * scale,
        y: y - blurred.padding * PX_TO_INCH * scale,
        w: w + blurred.padding * 2 * PX_TO_INCH * scale,
        h: h + blurred.padding * 2 * PX_TO_INCH * scale,
        data: blurred.data,
      },
      zIndex,
      domOrder: domOrder - 0.1,
    });
  }

  if (isTextContainer(node) && !clipped) {
    const padding = getPadding(style, scale);
    const textParts = collectTextParts(node, style, scale, null, true, opacity);
    const textOptions = {
      x: x + padding[3],
      y: y + padding[0],
      w: Math.max(0.01, w - padding[1] - padding[3]),
      h: Math.max(0.01, h - padding[0] - padding[2]),
      margin: 0,
      breakLine: false,
      rotate,
      fit: 'shrink',
      valign: style.alignItems === 'center' || style.justifyContent === 'center' ? 'mid' : 'top',
      vert: getWritingModeVert(style.writingMode, style.textOrientation),
      ...getTextStyle(style, scale, true, opacity),
    };
    items.push({ type: 'text', textParts, options: textOptions, zIndex, domOrder: domOrder + 0.2 });
    return { items, stopRecursion: true };
  }

  if (clipped) {
    const item = { type: 'image', options: { ...base }, zIndex, domOrder };
    return {
      items: [item],
      job: async () => {
        item.options.data = await elementToCanvasImage(node, rect.width, rect.height);
        if (!item.options.data) item.skip = true;
      },
      stopRecursion: true,
    };
  }

  return items.length ? { items } : null;
}
