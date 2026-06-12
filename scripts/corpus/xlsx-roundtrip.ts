// xlsx-writer round-trip gate (E-SHEET SD3, tail TD3c): for every xlsx in the
// corpus,
//   readXlsxToSheetDoc(bytes) → S0 → writeXlsx(S0) → readXlsxToSheetDoc(out) → S1
// and compare the SheetDoc signatures. The writer round-trips the grid surface
// (SD2 proves byte-stable identity on synthetic fixtures); this runs it against
// real-world files to confirm it neither crashes nor drops modelled data.
//
// Pure JS parse + write of our own bytes — no LibreOffice, no rendering, so no
// Docker sandbox is needed (the reader is already zip-bomb + DoS hardened for the
// untrusted corpus). Reports a markdown table to stdout. Embedded charts (a
// sheet's drawing) are not written back, but they live outside the grid surface,
// so a chart-bearing sheet still round-trips its cells/styles/CF/etc.
//
// Usage: CORPUS_DIR=corpus/external/poi-xlsx npx tsx scripts/corpus/xlsx-roundtrip.ts

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { basename, dirname, resolve } from 'node:path';

import { listCorpus } from './lib';
import type { SheetDoc } from '@/core/ir/sheet';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const corpusDir = process.env.CORPUS_DIR
  ? resolve(root, process.env.CORPUS_DIR)
  : resolve(root, 'corpus/external/poi-xlsx');

interface Signature {
  readonly sheets: number;
  readonly cells: number;
  readonly text: string;
  readonly sharedStrings: number;
  readonly merges: number;
  readonly cf: number;
  readonly sparklines: number;
  readonly tables: number;
  readonly styles: number;
}

function signature(d: SheetDoc): Signature {
  let cells = 0;
  let merges = 0;
  let cf = 0;
  let sparklines = 0;
  let tables = 0;
  let text = '';
  for (const s of d.sheets) {
    const g = s.grid;
    cells += g.cells.length;
    merges += g.merges.length;
    cf += (g.conditionalFormats ?? []).reduce((n, c) => n + c.rules.length, 0);
    sparklines += g.sparklines?.length ?? 0;
    tables += g.tables?.length ?? 0;
    for (const c of g.cells) text += c.inlineText ?? c.rawValue;
  }
  return {
    sheets: d.sheets.length,
    cells,
    text: text.replace(/\s+/g, ''),
    sharedStrings: d.sharedStrings.length,
    merges,
    cf,
    sparklines,
    tables,
    styles:
      d.styles.fonts.length +
      d.styles.fills.length +
      d.styles.borders.length +
      d.styles.cellXfs.length,
  };
}

function diffNote(a: Signature, b: Signature): string {
  const parts: Array<string> = [];
  if (a.text !== b.text) {
    const lost = a.text.length - b.text.length;
    parts.push(
      `text ${a.text.length}→${b.text.length} (${lost >= 0 ? '-' : '+'}${Math.abs(lost)})`,
    );
  }
  for (const k of [
    'sheets',
    'cells',
    'sharedStrings',
    'merges',
    'cf',
    'sparklines',
    'tables',
    'styles',
  ] as const) {
    if (a[k] !== b[k]) parts.push(`${k} ${a[k]}→${b[k]}`);
  }
  return parts.join(', ');
}

interface Row {
  // ✅ identical · ⚠️ divergent · ❌ writer threw · ⏭ input unreadable (reader)
  readonly name: string;
  readonly status: '✅' | '⚠️' | '❌' | '⏭';
  readonly note: string;
}

function main(): void {
  const inputs = listCorpus(corpusDir).filter((f) => /\.xlsx?$|\.xlsm$/i.test(f));
  if (inputs.length === 0) {
    console.error(`No xlsx in ${corpusDir}`);
    process.exit(1);
  }
  const rows: Array<Row> = [];
  let ok = 0;
  let warn = 0;
  let fail = 0;
  let skip = 0;

  for (const input of inputs) {
    const name = basename(input);
    // Reader rejections (encrypted, invalid zip, …) are NOT writer failures —
    // the writer never sees those workbooks. Separate the phases.
    let s0: SheetDoc;
    try {
      s0 = readXlsxToSheetDoc(new Uint8Array(readFileSync(input)));
    } catch {
      rows.push({ name, status: '⏭', note: 'input unreadable (reader)' });
      skip++;
      continue;
    }
    try {
      const written = writeXlsx(s0).bytes;
      const s1 = readXlsxToSheetDoc(written);
      const a = signature(s0);
      const b = signature(s1);
      const note = diffNote(a, b);
      if (note === '') {
        rows.push({ name, status: '✅', note: `${a.sheets}sh ${a.cells}c ${a.tables}t` });
        ok++;
      } else {
        rows.push({ name, status: '⚠️', note });
        warn++;
      }
    } catch (err) {
      rows.push({ name, status: '❌', note: (err as Error).message.slice(0, 70) });
      fail++;
    }
  }

  rows.sort((x, y) => (x.status + x.name).localeCompare(y.status + y.name));
  const readable = ok + warn + fail;
  const lines: Array<string> = [];
  lines.push(`# xlsx-writer round-trip — ${basename(corpusDir)}`);
  lines.push('');
  lines.push(
    `**${readable} readable books — ✅ ${ok} identical · ⚠️ ${warn} divergent · ` +
      `❌ ${fail} writer-failed; ⏭ ${skip} unreadable input.**`,
  );
  lines.push('');
  lines.push('| Book | St | Note |');
  lines.push('|---|---|---|');
  for (const r of rows) lines.push(`| ${r.name} | ${r.status} | ${r.note} |`);
  console.log(lines.join('\n'));
}

main();
