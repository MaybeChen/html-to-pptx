#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'fs/promises'
import { dirname, resolve } from 'path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const projectRoot = resolve(__dirname, '..')
const targetDir = resolve(projectRoot, 'dom-to-pptx')

export const DOM_TO_PPTX_UPSTREAM_BASE = 'https://raw.githubusercontent.com/atharva9167j/dom-to-pptx/master/src'

export const DOM_TO_PPTX_UPSTREAM_FILES = Object.freeze([
  'index.js',
  'font-embedder.js',
  'font-utils.js',
  'image-processor.js',
  'pptx-normalizer.js',
  'utils.js',
])

export function buildDomToPptxUpstreamUrl(fileName) {
  return `${DOM_TO_PPTX_UPSTREAM_BASE}/${fileName}`
}

export async function fetchDomToPptxSource(fileName, fetchImpl = globalThis.fetch) {
  const url = buildDomToPptxUpstreamUrl(fileName)
  const response = await fetchImpl(url)
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: HTTP ${response.status}`)
  }
  return response.text()
}

export async function readLocalDomToPptxSource(fileName, outputDir = targetDir) {
  return readFile(resolve(outputDir, fileName), 'utf8')
}

export async function collectDomToPptxSourceDiffs(options = {}) {
  const fetchImpl = options.fetch || globalThis.fetch
  const outputDir = options.outputDir || targetDir
  const diffs = []

  for (const fileName of DOM_TO_PPTX_UPSTREAM_FILES) {
    const [remoteSource, localSource] = await Promise.all([
      fetchDomToPptxSource(fileName, fetchImpl),
      readLocalDomToPptxSource(fileName, outputDir),
    ])

    if (localSource !== remoteSource) {
      diffs.push({
        fileName,
        url: buildDomToPptxUpstreamUrl(fileName),
        localBytes: Buffer.byteLength(localSource),
        upstreamBytes: Buffer.byteLength(remoteSource),
      })
    }
  }

  return diffs
}

export async function syncDomToPptxSources(options = {}) {
  const fetchImpl = options.fetch || globalThis.fetch
  const outputDir = options.outputDir || targetDir
  await mkdir(outputDir, { recursive: true })

  const synced = []
  for (const fileName of DOM_TO_PPTX_UPSTREAM_FILES) {
    const source = await fetchDomToPptxSource(fileName, fetchImpl)
    const outputPath = resolve(outputDir, fileName)
    await writeFile(outputPath, source, 'utf8')
    synced.push({ fileName, outputPath, url: buildDomToPptxUpstreamUrl(fileName) })
  }

  return synced
}

async function main() {
  if (process.argv.includes('--check')) {
    const diffs = await collectDomToPptxSourceDiffs()
    if (diffs.length === 0) {
      console.log('src/dom-to-pptx matches upstream exactly.')
      return
    }

    for (const diff of diffs) {
      console.error(
        `${diff.fileName} differs from ${diff.url} ` +
          `(local ${diff.localBytes} bytes, upstream ${diff.upstreamBytes} bytes)`
      )
    }
    process.exit(1)
  }

  const synced = await syncDomToPptxSources()
  for (const item of synced) {
    console.log(`synced ${item.fileName} <- ${item.url}`)
  }
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  main().catch((error) => {
    console.error(error)
    process.exit(1)
  })
}
