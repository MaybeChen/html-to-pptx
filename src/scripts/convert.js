#!/usr/bin/env node

import { access, readFile, rm, stat, writeFile } from 'fs/promises'
import { basename, dirname, isAbsolute, join, parse, relative, resolve } from 'path'
import { fileURLToPath, pathToFileURL } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const projectRoot = resolve(__dirname, '..')
const defaultAssetsDir = resolve(projectRoot, 'assets')
const knownAssetMirrors = [
  {
    remotePrefix: 'https://cdn.digitalhumanai.top/slidagent/pptx-craft/assets/',
    localDir: defaultAssetsDir,
  },
  {
    remotePrefix: 'https://npmmirror.com/mirrors/fonteditor-core@2.6.3/',
    localDir: join(defaultAssetsDir, 'fonteditor-core@2.6.3'),
  },
]
const defaultEmbedFonts = [
  {
    name: 'Noto Sans SC',
    url: 'https://cdn.digitalhumanai.top/slidagent/pptx-craft/assets/fonts/NotoSansSC/NotoSansSC-Regular.ttf',
  },
]
let sharedBrowser = null

function toPosixPath(value) {
  return String(value).replace(/\\/g, '/')
}

function longestCommonDir(paths) {
  if (!Array.isArray(paths) || paths.length === 0) {
    throw new Error('无法从空输入中解析 htmlDir')
  }
  const dirs = paths.map((item) => resolve(dirname(item)))
  if (dirs.length === 1) return dirs[0]
  const splitDirs = dirs.map((item) => toPosixPath(item).split('/'))
  let common = splitDirs[0]
  for (const parts of splitDirs.slice(1)) {
    let idx = 0
    while (idx < common.length && idx < parts.length && common[idx].toLowerCase() === parts[idx].toLowerCase()) idx += 1
    common = common.slice(0, idx)
  }
  const root = parse(dirs[0]).root
  if (!common.length) return root
  return common.join('/') || root
}

export function resolveHtmlDirFromInput(input) {
  if (Array.isArray(input)) return longestCommonDir(input)
  const absolute = resolve(input)
  return /\.html?$/i.test(absolute) ? dirname(absolute) : absolute
}

export function buildRenderPageUrl(baseUrl, serverRoot, htmlPath) {
  const root = resolve(serverRoot)
  const absoluteHtmlPath = resolve(htmlPath)
  const rel = relative(root, absoluteHtmlPath)
  if (!(rel === '' || (!rel.startsWith('..') && !isAbsolute(rel)))) {
    throw new Error(`HTML 文件不在服务目录下: ${absoluteHtmlPath}`)
  }
  const encodedPath = toPosixPath(rel)
    .replace(/^\/+/, '')
    .split('/')
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join('/')
  return new URL(encodedPath || '.', `${baseUrl.replace(/\/$/, '')}/`).toString()
}

async function fileExists(path) {
  try {
    await access(path)
    return true
  } catch {
    return false
  }
}

export function mapKnownAssetUrlToLocalUrl(url, baseUrl = '') {
  for (const mirror of knownAssetMirrors) {
    if (!url.startsWith(mirror.remotePrefix)) continue
    const rel = decodeURIComponent(url.slice(mirror.remotePrefix.length).split(/[?#]/)[0])
    const localPath = resolve(mirror.localDir, rel)
    const routePath = `/__local_asset__/${encodeURIComponent(toPosixPath(relative(defaultAssetsDir, localPath)))}`
    return baseUrl ? new URL(routePath, baseUrl).toString() : routePath
  }
  return url
}

export function rewriteHtmlCdnUrlsToLocalAssets(html, baseUrl = '') {
  let result = String(html)
  for (const mirror of knownAssetMirrors) {
    result = result.split(mirror.remotePrefix).join(mapKnownAssetUrlToLocalUrl(mirror.remotePrefix, baseUrl).replace(/\/[^/]*$/, '/'))
  }
  return result
}

function rewritePageAssetUrlsToLocalAssets(html, baseUrl) {
  return String(html).replace(/https?:\/\/[^'"\s)]+/g, (url) => mapKnownAssetUrlToLocalUrl(url, baseUrl))
}

function resolveAssetUrl(rawUrl, pageUrl) {
  if (!rawUrl || /^(data:|blob:|mailto:|tel:|#)/i.test(rawUrl)) return rawUrl
  return new URL(rawUrl, pageUrl).toString()
}

function absolutizeCssUrls(cssText, pageUrl) {
  return String(cssText).replace(/url\((['"]?)(.*?)\1\)/gi, (full, quote, rawUrl) => {
    const trimmed = String(rawUrl || '').trim()
    if (!trimmed) return full
    return `url(${quote}${resolveAssetUrl(trimmed, pageUrl)}${quote})`
  })
}

function absolutizeSrcset(srcset, pageUrl) {
  return String(srcset)
    .split(',')
    .map((candidate) => {
      const parts = candidate.trim().split(/\s+/)
      if (!parts[0]) return candidate
      parts[0] = resolveAssetUrl(parts[0], pageUrl)
      return parts.join(' ')
    })
    .join(', ')
}

function absolutizeHtmlAssetUrls(html, pageUrl) {
  let result = String(html)
  result = result.replace(/\b(src|href|poster)\s*=\s*(['"])(.*?)\2/gi, (full, attr, quote, value) => {
    return `${attr}=${quote}${resolveAssetUrl(value, pageUrl)}${quote}`
  })
  result = result.replace(/\bsrcset\s*=\s*(['"])(.*?)\1/gi, (full, quote, value) => {
    return `srcset=${quote}${absolutizeSrcset(value, pageUrl)}${quote}`
  })
  result = result.replace(/<style\b[^>]*>([\s\S]*?)<\/style>/gi, (full, css) => {
    return full.replace(css, absolutizeCssUrls(css, pageUrl))
  })
  result = result.replace(/style\s*=\s*(['"])(.*?)\1/gi, (full, quote, css) => {
    return `style=${quote}${absolutizeCssUrls(css, pageUrl)}${quote}`
  })
  return result
}

export function absolutizeExtractedPageAssets(pageSnapshot, pageUrl, baseUrl = '') {
  return rewritePageAssetUrlsToLocalAssets(absolutizeHtmlAssetUrls(pageSnapshot, pageUrl), baseUrl)
}

export async function startRenderServer(htmlDir, options = {}) {
  const { default: express } = await import('express')
  const { default: getPort } = await import('get-port')
  const root = resolve(htmlDir)
  const port = options.port || (await getPort({ port: [4173, 5173, 6173] }))
  const app = express()
  app.get('/__health', (_req, res) => res.type('text/plain').send('ok'))
  app.get('/__local_asset__/:assetPath', async (req, res) => {
    const assetPath = resolve(defaultAssetsDir, decodeURIComponent(req.params.assetPath))
    const rel = relative(defaultAssetsDir, assetPath)
    if (rel.startsWith('..') || isAbsolute(rel) || !(await fileExists(assetPath))) {
      res.status(404).send('Not found')
      return
    }
    res.sendFile(assetPath)
  })
  app.use(express.static(root, { fallthrough: false }))
  const server = await new Promise((resolveServer) => {
    const listener = app.listen(port, '127.0.0.1', () => resolveServer(listener))
  })
  const baseUrl = `http://127.0.0.1:${port}/`
  return {
    app,
    server,
    root,
    port,
    baseUrl,
    close: () => new Promise((resolveClose, rejectClose) => server.close((error) => (error ? rejectClose(error) : resolveClose()))),
  }
}

export async function removeTemporaryFile(path) {
  if (!path) return
  await rm(path, { force: true })
}

export function buildTempHtmlPath(inputHtmlPath, suffix = '.localized.html') {
  const parsed = parse(resolve(inputHtmlPath))
  return join(parsed.dir, `${parsed.name}${suffix}`)
}

export function buildExportOptions(options = {}) {
  return {
    fileName: options.fileName || 'output.pptx',
    autoEmbedFonts: options.autoEmbedFonts ?? true,
    svgAsVector: options.svgAsVector ?? false,
    layout: options.layout || 'LAYOUT_WIDE',
    skipDownload: options.skipDownload ?? false,
    ...options,
  }
}

export function buildEmbedFontsConfig(fonts = defaultEmbedFonts) {
  return fonts.map((font) => ({ ...font }))
}

export function patchBundleSourceForLocalAssets(source, baseUrl = '') {
  return rewritePageAssetUrlsToLocalAssets(source, baseUrl)
}

export function shouldIgnoreBrowserConsoleMessage(message = '') {
  const text = typeof message === 'string' ? message : message.text?.() || String(message)
  return [
    'Failed to load resource: net::ERR_ABORTED',
    'favicon.ico',
    'ResizeObserver loop',
    'Download the Vue Devtools extension',
  ].some((item) => text.includes(item))
}

export function attachPageDiagnostics(page, logger = console) {
  page.on('console', (message) => {
    const text = message.text()
    if (!shouldIgnoreBrowserConsoleMessage(text)) logger.log(`[browser:${message.type()}] ${text}`)
  })
  page.on('pageerror', (error) => logger.error('[browser:pageerror]', error))
  page.on('requestfailed', (request) => logger.warn('[browser:requestfailed]', request.url(), request.failure()?.errorText))
}

async function createLocalizedHtmlCopy(inputHtmlPath, baseUrl) {
  const source = await readFile(inputHtmlPath, 'utf8')
  const localized = rewriteHtmlCdnUrlsToLocalAssets(source, baseUrl)
  if (localized === source) return inputHtmlPath
  const tempPath = buildTempHtmlPath(inputHtmlPath)
  await writeFile(tempPath, localized)
  return tempPath
}

function collectTargetsScript() {
  return `(() => {
    const selectors = ['.slide', '#slide', '.slide-container'];
    for (const selector of selectors) {
      const nodes = [...document.querySelectorAll(selector)];
      if (nodes.length) return nodes;
    }
    return [document.body || document.documentElement];
  })()`
}

async function waitForPageStable(page) {
  await page.waitForLoadState('domcontentloaded')
  await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {})
  await page.evaluate(async () => {
    if (document.fonts?.ready) await document.fonts.ready
    await Promise.all([...document.images].map((img) => img.complete ? undefined : new Promise((resolve) => {
      img.addEventListener('load', resolve, { once: true })
      img.addEventListener('error', resolve, { once: true })
    })))
    await new Promise((resolve) => setTimeout(resolve, 300))
  })
}

async function runBrowserExport(page, outputPath, options) {
  const exportOptions = buildExportOptions({ ...options, fileName: basename(outputPath), skipDownload: false })
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: options.timeout ?? 120_000 }),
    page.evaluate(
      async ({ script, exportOptions }) => {
        const { exportToPptx } = await import('https://esm.sh/dom-to-pptx')
        const targets = eval(script)
        await exportToPptx(targets, exportOptions)
      },
      { script: collectTargetsScript(), exportOptions },
    ),
  ])
  await download.saveAs(outputPath)
}

export async function convertHtmlToPptx(input, outputPath, options = {}) {
  const htmlFiles = Array.isArray(input) ? input.map((item) => resolve(item)) : [resolve(input)]
  const htmlDir = options.htmlDir ? resolve(options.htmlDir) : resolveHtmlDirFromInput(htmlFiles)
  const server = await startRenderServer(htmlDir, options.server || {})
  const { chromium } = await import('playwright')
  let browser = options.browser || sharedBrowser
  let ownsBrowser = false
  if (!browser) {
    browser = await chromium.launch({ headless: true, args: ['--allow-file-access-from-files', '--disable-web-security'] })
    ownsBrowser = true
  }
  const context = await browser.newContext({ acceptDownloads: true, viewport: options.viewport || { width: 1920, height: 1080 } })
  const page = await context.newPage()
  attachPageDiagnostics(page, options.logger || console)
  try {
    for (const htmlFile of htmlFiles) {
      const localizedHtmlPath = await createLocalizedHtmlCopy(htmlFile, server.baseUrl)
      const pageUrl = buildRenderPageUrl(server.baseUrl, htmlDir, localizedHtmlPath)
      await page.goto(pageUrl, { waitUntil: 'domcontentloaded', timeout: options.timeout ?? 60_000 })
      await waitForPageStable(page)
      await runBrowserExport(page, outputPath, options)
      if (localizedHtmlPath !== htmlFile) await removeTemporaryFile(localizedHtmlPath)
    }
    const info = await stat(outputPath)
    return { outputPath, size: info.size }
  } finally {
    await context.close().catch(() => {})
    if (ownsBrowser) await browser.close().catch(() => {})
    await server.close().catch(() => {})
  }
}

export async function closeBrowser() {
  if (!sharedBrowser) return
  await sharedBrowser.close()
  sharedBrowser = null
}

async function convertSingleFile(inputPath, outputPath, options = {}) {
  console.log(` 转换: ${inputPath}`)
  return convertHtmlToPptx(inputPath, outputPath, options)
}

async function main() {
  const args = process.argv.slice(2)
  const input = args[0]
  const output = args[1] || (input ? input.replace(/\.pptx\.html$/i, '.pptx').replace(/\.html?$/i, '.pptx') : '')

  console.log('========================================')
  console.log(' HTML 转 PPTX')
  console.log('========================================')

  if (!input) {
    console.error('用法: node src/convert.js <input.html> [output.pptx]')
    process.exit(1)
  }

  try {
    console.log(`输入: ${input}`)
    console.log(`输出: ${output}`)
    console.log(' 开始转换...')
    const result = await convertSingleFile(input, output)
    console.log('✅ 转换完成')
    console.log(`输出文件: ${result.outputPath}`)
    console.log(`文件大小: ${(result.size / 1024).toFixed(2)} KB`)
  } catch (error) {
    console.error('❌ 转换失败')
    console.error('错误:', error?.message || error)
    if (error?.stack) console.error(error.stack)
    process.exit(1)
  } finally {
    await closeBrowser()
  }
}

if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  main().catch((error) => {
    console.error('执行失败:', error)
    process.exit(1)
  })
}
