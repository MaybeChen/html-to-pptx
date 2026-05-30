import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';

import { buildDomToPptxModuleUrl, buildExportOptions, collectHtmlFiles, createMergedHtmlFile, startRenderServer } from '../src/scripts/convert.js';
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

test('collectHtmlFiles returns sorted top-level html files', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'html-to-pptx-'));
  try {
    await writeFile(join(dir, '10.html'), '<html></html>');
    await writeFile(join(dir, '2.htm'), '<html></html>');
    await writeFile(join(dir, 'ignore.txt'), 'ignore');

    const files = await collectHtmlFiles(dir);

    assert.deepEqual(files.map((file) => file.split(/[\\/]/).pop()), ['2.htm', '10.html']);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('createMergedHtmlFile wraps each html file as a ppt slide', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'html-to-pptx-'));
  try {
    const first = join(dir, 'a.html');
    const second = join(dir, 'b.html');
    await writeFile(first, '<html><head><link rel="stylesheet" href="./a.css"></head><body><h1>A</h1></body></html>');
    await writeFile(second, '<html><head><style>.b{color:red}</style></head><body class="source"><h1>B</h1></body></html>');

    const { tempPath } = await createMergedHtmlFile(dir, [first, second], 'http://127.0.0.1:4173/');
    const merged = await readFile(tempPath, 'utf8');

    assert.match(merged, /class="ppt-slide"/);
    assert.match(merged, /data-source="a\.html"/);
    assert.match(merged, /data-source="b\.html"/);
    assert.match(merged, /href="http:\/\/127\.0\.0\.1:4173\/a\.css"/);
    assert.match(merged, new RegExp(DEFAULT_FONT_CSS_URLS[0].replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('merged html file is served by the render server', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'html-to-pptx-'));
  let server;
  try {
    const first = join(dir, 'a.html');
    await writeFile(first, '<html><body><h1>A</h1></body></html>');

    server = await startRenderServer(dir);
    const { tempPath } = await createMergedHtmlFile(dir, [first], server.baseUrl);

    assert.equal(basename(tempPath).startsWith('.'), false);

    const response = await fetch(new URL(basename(tempPath), server.baseUrl));
    assert.equal(response.status, 200);
    assert.match(await response.text(), /class="ppt-slide"/);
  } finally {
    if (server) await server.close().catch(() => {});
    await rm(dir, { recursive: true, force: true });
  }
});

test('render server serves the local dom-to-pptx browser module', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'html-to-pptx-'));
  let server;
  try {
    server = await startRenderServer(dir);
    const response = await fetch(buildDomToPptxModuleUrl(server.baseUrl));
    const source = await response.text();

    assert.equal(response.status, 200);
    assert.match(source, /export async function exportToPptx/);
    assert.doesNotMatch(source, /https:\/\/esm\.sh\/dom-to-pptx/);
    assert.match(source, /\/__dom_to_pptx_vendor__\/pptxgenjs\.js/);
  } finally {
    if (server) await server.close().catch(() => {});
    await rm(dir, { recursive: true, force: true });
  }
});

test('render server does not read vendor directories as files', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'html-to-pptx-'));
  let server;
  try {
    server = await startRenderServer(dir);
    const response = await fetch(new URL('/__dom_to_pptx_vendor__/fonteditor-core/src/ttf/table', server.baseUrl));

    assert.equal(response.status, 404);
  } finally {
    if (server) await server.close().catch(() => {});
    await rm(dir, { recursive: true, force: true });
  }
});
