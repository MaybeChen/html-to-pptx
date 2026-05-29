import { Font } from 'fonteditor-core';
import pako from 'pako';

export async function fontToEot(type, fontBuffer) {
  const font = Font.create(fontBuffer, {
    type,
    hinting: true,
    inflate: type === 'woff' ? pako.inflate : undefined,
  });
  const eotBuffer = font.write({ type: 'eot', toBuffer: true });
  if (eotBuffer instanceof ArrayBuffer) return eotBuffer;
  return eotBuffer.buffer.slice(eotBuffer.byteOffset, eotBuffer.byteOffset + eotBuffer.byteLength);
}
