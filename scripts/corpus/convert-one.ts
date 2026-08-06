// Convert ONE document to PDF in an isolated child process. The corpus runner
// uses this in sandbox mode so a hostile or pathological input can't hang or
// OOM the parent: the parent spawns this with a wall-clock timeout and a heap
// cap (NODE_OPTIONS=--max-old-space-size). Our OpcPackage.open already caps
// decompression; this adds CPU/time isolation on top.
//
// Usage: tsx convert-one.ts <input> <outPdf>

import { readFileSync, writeFileSync } from 'node:fs';

import { corpusFontOptions } from './fonts';
import { Ream } from '@/core/converter/ream';

const [input, outPdf] = process.argv.slice(2);
if (!input || !outPdf) {
  console.error('usage: convert-one <input> <outPdf>');
  process.exit(2);
}

// The Ream facade sniffs the format and dispatches — one path for every input
// (docx/xlsx/pptx/pdf + legacy doc/xls/ppt), which is what makes this child a
// safe universal isolator for untrusted corpus files.
const bytes = new Uint8Array(readFileSync(input));
// The substitutes come off disk (corpus/.fonts), so a child process pays no
// download: the cache the parent fills is the one this reads.
const pdf = await Ream.parse(bytes).convert('pdf', corpusFontOptions());
writeFileSync(outPdf, pdf);
