// Convert ONE document to PDF in an isolated child process. The corpus runner
// uses this in sandbox mode so a hostile or pathological input can't hang or
// OOM the parent: the parent spawns this with a wall-clock timeout and a heap
// cap (NODE_OPTIONS=--max-old-space-size). Our OpcPackage.open already caps
// decompression; this adds CPU/time isolation on top.
//
// Usage: tsx convert-one.ts <input> <outPdf>

import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { FontBytesByVariant } from '@/font';
import { convertDocxToPdfSync, convertXlsxToPdfSync } from '@/converter';

const [input, outPdf] = process.argv.slice(2);
if (!input || !outPdf) {
  console.error('usage: convert-one <input> <outPdf>');
  process.exit(2);
}

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const font = (n: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts', n)));
const fonts: FontBytesByVariant = {
  regular: font('Roboto-Regular.ttf'),
  bold: font('Roboto-Bold.ttf'),
  italic: font('Roboto-Italic.ttf'),
  boldItalic: font('Roboto-BoldItalic.ttf'),
};

const bytes = new Uint8Array(readFileSync(input));
const pdf = /\.(xlsx|xlsm)$/i.test(input)
  ? convertXlsxToPdfSync(bytes, { fonts })
  : convertDocxToPdfSync(bytes, { fonts });
writeFileSync(outPdf, pdf);
