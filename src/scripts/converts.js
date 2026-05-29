#!/usr/bin/env node

import { mkdir, readdir, readFile, rm, stat, writeFile } from 'fs/promises'
import { basename, extname, isAbsolute, join, relative, resolve } from 'path'
import { pathToFileURL } from 'url'

import {
  attachPageDiagnostics,
  buildExportOptions,
  buildRenderPageUrl,
  startRenderServer,
} from './convert.js'
import { collectMergedStylesheetHrefs } from './merge-html-assets.js'
import { waitForFontsReady } from './wait-for-fonts.js'

const DEFAULT_HTML_EXTENSIONS = new Set(['.html', '.htm'])
const MERGED_HTML_FILE_NAME = '.__html_to_pptx_merged__.html'

function toPosixPath(value) {
  return String(value).replace(/\\/g, '/')
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function normalizeOutputPath(outputPath) {
  const resolved = resolve(outputPath)
  return resolved.toLowerCase().endsWith('.pptx') ? resolved : `${resolved}.pptx`
}

export async function collectHtmlFiles(inputDir, options = {}) {
  const root = resolve(inputDir)
  const recursive = options.recursive ?? false
  const entries = await readdir(root, { withFileTypes: true })
  const files = []

  for (const entry of entries) {
    const fullPath = join(root, entry.name)
    if (entry.isDirectory()) {
      if (recursive) files.push(...await collectHtmlFiles(fullPath, options))
      continue
    }

    if (!entry.isFile()) continue
    if (entry.name === MERGED_HTML_FILE_NAME) continue
    if (DEFAULT_HTML_EXTENSIONS.has(extname(entry.name).toLowerCase())) files.push(fullPath)
  }

  return files.sort((a, b) => toPosixPath(relative(root, a)).localeCompare(toPosixPath(relative(root, b)), undefined, { numeric: true }))
}

function extractTagContent(html, tagName) {
  const match = String(html).match(new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)<\\/${tagName}>`, 'i'))
  return match ? match[1] : ''
}

function extractTagAttributes(html, tagName) {
  const match = String(html).match(new RegExp(`<${tagName}\\b([^>]*)>`, 'i'))
  return match ? match[1].trim() : ''
}

function stripScripts(html) {
  return String(html).replace(/<script\b[\s\S]*?<\/script>/gi, '')
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

function absolutizeHtmlAssets(html, pageUrl) {
  let result = String(html)
  result = result.replace(/\b(src|href|poster)\s*=\s*(['"])(.*?)\2/gi, (_full, attr, quote, value) => {
    return `${attr}=${quote}${resolveAssetUrl(value, pageUrl)}${quote}`
  })
  result = result.replace(/\bsrcset\s*=\s*(['"])(.*?)\1/gi, (_full, quote, value) => {
    return `srcset=${quote}${absolutizeSrcset(value, pageUrl)}${quote}`
  })
  result = result.replace(/<style\b[^>]*>([\s\S]*?)<\/style>/gi, (full, css) => {
    return full.replace(css, absolutizeCssUrls(css, pageUrl))
  })
  result = result.replace(/style\s*=\s*(['"])(.*?)\1/gi, (_full, quote, css) => {
    return `style=${quote}${absolutizeCssUrls(css, pageUrl)}${quote}`
  })
  return result
}

function collectExternalStylesheetHrefs(headHtml) {
  const hrefs = []
  String(headHtml).replace(/<link\b[^>]*>/gi, (tag) => {
    if (!/rel\s*=\s*(['"]?)stylesheet\1/i.test(tag)) return tag
    const hrefMatch = tag.match(/href\s*=\s*(['"])(.*?)\1/i)
    if (hrefMatch?.[2]) hrefs.push(hrefMatch[2])
    return tag
  })
  return hrefs
}

function collectInlineStyles(headHtml) {
  const styles = []
  String(headHtml).replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, (tag) => {
    styles.push(tag)
    return tag
  })
  return styles
}

export async function createMergedHtmlFile(inputDir, htmlFiles, serverBaseUrl) {
  const items = []

  for (const htmlFile of htmlFiles) {
    const pageUrl = buildRenderPageUrl(serverBaseUrl, inputDir, htmlFile)
    const html = await readFile(htmlFile, 'utf8')
    const head = absolutizeHtmlAssets(extractTagContent(html, 'head'), pageUrl)
    const body = absolutizeHtmlAssets(extractTagContent(html, 'body') || html, pageUrl)
    const bodyAttrs = extractTagAttributes(html, 'body')

    items.push({
      sourcePath: htmlFile,
      pageUrl,
      externalLinks: collectExternalStylesheetHrefs(head),
      inlineStyles: collectInlineStyles(head),
      bodyAttrs,
      bodyHtml: stripScripts(body),
    })
  }

  const stylesheetLinks = collectMergedStylesheetHrefs(items)
    .map((href) => `    <link rel="stylesheet" href="${escapeHtml(href)}">`)
    .join('\n')
  const inlineStyles = items.flatMap((item) => item.inlineStyles).join('\n')
  const slides = items
    .map((item, index) => {
      const relativeSource = toPosixPath(relative(inputDir, item.sourcePath))
      const sourceAttrs = item.bodyAttrs ? ` ${item.bodyAttrs}` : ''
      return `    <section class="ppt-slide" data-source="${escapeHtml(relativeSource)}" data-slide-index="${index + 1}"${sourceAttrs}>\n${item.bodyHtml}\n    </section>`
    })
    .join('\n')

  const mergedHtml = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Merged HTML to PPTX</title>
${stylesheetLinks}
${inlineStyles}
  <style>
    html, body { margin: 0; padding: 0; background: transparent; }
    .ppt-slide { position: relative; display: block; box-sizing: border-box; overflow: hidden; }
  </style>
</head>
<body>
${slides}
</body>
</html>
`

  const tempPath = join(resolve(inputDir), MERGED_HTML_FILE_NAME)
  await writeFile(tempPath, mergedHtml, 'utf8')
  return { tempPath, items }
}

async function waitForPageStable(page) {
  await page.waitForLoadState('domcontentloaded')
  await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {})
  await waitForFontsReady(page, { selector: '.ppt-slide' })
  await page.evaluate(async () => {
    await Promise.all([...document.images].map((img) => img.complete ? undefined : new Promise((resolve) => {
      img.addEventListener('load', resolve, { once: true })
      img.addEventListener('error', resolve, { once: true })
    })))
  })
}

async function exportMergedPage(page, outputPath, options = {}) {
  const exportOptions = buildExportOptions({
    ...options,
    fileName: basename(outputPath),
    skipDownload: false,
  })

  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: options.timeout ?? 120_000 }),
    page.evaluate(async ({ exportOptions }) => {
      const { exportToPptx } = await import('https://esm.sh/dom-to-pptx')
      const targets = Array.from(document.querySelectorAll('.ppt-slide'))
      await exportToPptx(targets, exportOptions)
    }, { exportOptions }),
  ])

  await download.saveAs(outputPath)
}

export async function convertHtmlDirectoryToPptx(inputDir, outputPath, options = {}) {
  const root = resolve(inputDir)
  const output = normalizeOutputPath(outputPath)
  const inputStat = await stat(root)
  if (!inputStat.isDirectory()) throw new Error(`输入路径不是目录: ${root}`)

  const htmlFiles = await collectHtmlFiles(root, options)
  if (htmlFiles.length === 0) throw new Error(`目录下没有找到 HTML 文件: ${root}`)

  await mkdir(resolve(output, '..'), { recursive: true })
  const server = await startRenderServer(root, options.server || {})
  const { chromium } = await import('playwright')
  let browser
  let context
  let merged

  try {
    merged = await createMergedHtmlFile(root, htmlFiles, server.baseUrl)
    browser = options.browser || await chromium.launch({ headless: true, args: ['--allow-file-access-from-files', '--disable-web-security'] })
    context = await browser.newContext({ acceptDownloads: true, viewport: options.viewport || { width: 1920, height: 1080 } })
    const page = await context.newPage()
    attachPageDiagnostics(page, options.logger || console)

    const mergedUrl = buildRenderPageUrl(server.baseUrl, root, merged.tempPath)
    await page.goto(mergedUrl, { waitUntil: 'domcontentloaded', timeout: options.timeout ?? 60_000 })
    await waitForPageStable(page)
    await exportMergedPage(page, output, options)

    const outputStat = await stat(output)
    return {
      inputDir: root,
      outputPath: output,
      size: outputStat.size,
      slideCount: htmlFiles.length,
      htmlFiles,
    }
  } finally {
    if (context) await context.close().catch(() => {})
    if (browser && !options.browser) await browser.close().catch(() => {})
    if (merged?.tempPath) await rm(merged.tempPath, { force: true }).catch(() => {})
    await server.close().catch(() => {})
  }
}

async function main() {
  const [inputDir, outputPath] = process.argv.slice(2)

  if (!inputDir || !outputPath) {
    console.error('用法: node ./src/scripts/converts.js <输入目录> <输出文件.pptx>')
    process.exit(1)
  }

  try {
    console.log('========================================')
    console.log(' HTML 目录转 PPTX')
    console.log('========================================')
    console.log(`输入目录: ${inputDir}`)
    console.log(`输出文件: ${outputPath}`)

    const result = await convertHtmlDirectoryToPptx(inputDir, outputPath)

    console.log('✅ 转换完成')
    console.log(`HTML 文件数: ${result.slideCount}`)
    console.log(`输出文件: ${result.outputPath}`)
    console.log(`文件大小: ${(result.size / 1024).toFixed(2)} KB`)
  } catch (error) {
    console.error('❌ 转换失败')
    console.error('错误:', error?.message || error)
    if (error?.stack) console.error(error.stack)
    process.exit(1)
  }
}

if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  main().catch((error) => {
    console.error('执行失败:', error)
    process.exit(1)
  })
}
