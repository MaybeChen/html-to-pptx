export const DEFAULT_FONT_CSS_URLS = [
  'https://cdn.digitalhumanai.top/slidagent/pptx-craft/assets/css/fonts.css'
];

export function collectMergedStylesheetHrefs(items) {
  const externalLinks = items.flatMap(item => item.externalLinks || []);

  return [...new Set([
    ...externalLinks,
    ...DEFAULT_FONT_CSS_URLS
  ])];
}
