#!/usr/bin/env node

import { convertHtmlDirectoryToPptx } from './index.js';

const [, , inputDir, outputFile] = process.argv;

if (!inputDir || !outputFile) {
  console.error('Usage: html-to-pptx <input-dir> <output-file>');
  process.exitCode = 1;
} else {
  await convertHtmlDirectoryToPptx({ inputDir, outputFile });
}
