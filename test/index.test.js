import assert from 'node:assert/strict';
import test from 'node:test';

import { convertHtmlDirectoryToPptx } from '../src/index.js';

test('convertHtmlDirectoryToPptx is a placeholder', async () => {
  await assert.rejects(
    () => convertHtmlDirectoryToPptx({ inputDir: 'input', outputFile: 'output.pptx' }),
    /not implemented yet/
  );
});
