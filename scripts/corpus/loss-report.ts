// What the reader actually SAYS about a corpus — and how many documents it
// hands back empty.
//
// The pixel scouts answer "how far is the page from its source". Neither
// answers the two questions a user asks first: did anything come out at all,
// and is the loss report telling me the truth. Both were being missed.
//
// This one found, on the four hundred files of the pdf.js corpus:
//   · 44 documents that reconstruct to NOTHING — a blank result, which every
//     pixel score reads as "a page that differs somewhat" rather than "empty";
//   · a `bare-shading (sh)` loss reported on all 400 files, most of which
//     contain no `sh` at all — a report that cries wolf on every document
//     tells a reader nothing.
//
// Usage:
//   npx tsx scripts/corpus/loss-report.ts [corpus-dir] [limit]

import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';

import { Ream } from '@/core/converter/ream';

/** How many example file names to keep beside each distinct loss. */
const EXAMPLES = 3;

/** How many distinct losses to print. */
const ROWS = 25;

/** A loss's text with the file-specific parts blurred, so like joins like. */
function shape(detail: string): string {
  return detail
    .replace(/\d+/gu, 'N')
    .replace(/[“"][^”"]*[”"]/gu, '…')
    .slice(0, 76);
}

function main(): void {
  const dir = process.argv[2] ?? 'corpus/external/pdfjs';
  const limit = Number(process.argv[3] ?? 10_000);
  const files = readdirSync(dir)
    .filter((n) => n.toLowerCase().endsWith('.pdf'))
    .slice(0, limit);
  const tally = new Map<string, { n: number; files: Array<string> }>();
  const note = (key: string, file: string): void => {
    const had = tally.get(key) ?? { n: 0, files: [] };
    had.n++;
    if (had.files.length < EXAMPLES) had.files.push(file);
    tally.set(key, had);
  };
  const empty: Array<string> = [];

  for (const f of files) {
    try {
      const doc = Ream.parse(new Uint8Array(readFileSync(resolve(dir, f))));
      if (doc.flow.body.length === 0) empty.push(f);
      for (const l of doc.losses) note(`${l.severity} | ${shape(l.detail)}`, f);
    } catch (e) {
      note(`THREW    | ${String(e).slice(0, 64)}`, f);
    }
  }

  process.stdout.write(`${String(files.length)} files\n\n`);
  process.stdout.write(`EMPTY — reconstructed to nothing: ${String(empty.length)}\n`);
  for (const f of empty) process.stdout.write(`  ${f}\n`);
  process.stdout.write('\n');
  for (const [key, v] of [...tally].sort((a, b) => b[1].n - a[1].n).slice(0, ROWS)) {
    process.stdout.write(
      `${String(v.n).padStart(4)}  ${key}\n        e.g. ${v.files.join(', ')}\n`,
    );
  }
}

main();
