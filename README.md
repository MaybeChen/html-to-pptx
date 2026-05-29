# html-to-pptx

A Node.js project scaffold for converting every HTML file in a directory into a complete PowerPoint presentation.

The conversion implementation is intentionally left as a placeholder until the technical details are finalized.

## Requirements

- Node.js 20 or newer

## Usage

```bash
npm start -- <input-dir> <output-file>
```

Or, after linking/installing the package:

```bash
html-to-pptx <input-dir> <output-file>
```

## Core conversion module

The vendored browser-side conversion core lives in `src/dom-to-pptx/src`. It exposes `exportToPptx(target, options)`, which accepts a DOM element, selector, or array of elements/selectors and generates one PPTX slide per root element.

This module is the integration point for the upcoming directory-level HTML-to-PPTX workflow.

## Development

```bash
npm test
```
