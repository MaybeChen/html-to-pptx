import assert from 'node:assert/strict';
import test from 'node:test';

import { buildExportOptions } from '../src/scripts/convert.js';
import { collectMergedStylesheetHrefs, DEFAULT_FONT_CSS_URLS } from '../src/scripts/merge-html-assets.js';

test('buildExportOptions applies defaults', () => {
  assert.deepEqual(buildExportOptions({ skipDownload: true }), {
    fileName: 'output.pptx',
    autoEmbedFonts: true,
    svgAsVector: false,
    layout: 'LAYOUT_WIDE',
    skipDownload: true,
  });
});

test('collectMergedStylesheetHrefs de-duplicates external and default font links', () => {
  assert.deepEqual(
    collectMergedStylesheetHrefs([
      { externalLinks: ['https://example.com/a.css', DEFAULT_FONT_CSS_URLS[0]] },
      { externalLinks: ['https://example.com/a.css', 'https://example.com/b.css'] },
    ]),
    ['https://example.com/a.css', DEFAULT_FONT_CSS_URLS[0], 'https://example.com/b.css']
  );
});
