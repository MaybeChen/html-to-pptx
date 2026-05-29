export async function getProcessedImage(node) {
  if (node.tagName === 'CANVAS') {
    return node.toDataURL('image/png');
  }

  if (node.tagName === 'SVG') {
    return `data:image/svg+xml;base64,${btoa(new XMLSerializer().serializeToString(node))}`;
  }

  if (node.currentSrc || node.src) {
    return node.currentSrc || node.src;
  }

  return undefined;
}
