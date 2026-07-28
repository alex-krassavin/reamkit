// Worker half of the xlsx invariant sweep (see xlsx-invariants.ts).
//
// Reads a chunk of .xlsx paths and appends one JSONL record per file to the
// output log — a `start` line before parsing and a `done` line after. The
// parent runs us with a heap cap and a wall-clock timeout, so a file that
// OOMs or hangs kills THIS process; the parent then reads the log, sees a
// `start` with no matching `done`, and attributes the crash to that file.
//
// Usage: tsx xlsx-invariant-child.ts <jsonl-out> <file>...

import { appendFileSync, readFileSync } from 'node:fs';

import type { BodyElement } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import { OpcPackage } from '@/core/opc';
import { readXlsx } from '@/excel/xlsx-reader';

/** One measurement of a single .xlsx — the unit the parent classifies. */
export interface FileProbe {
  /** Corpus-relative path, so records are stable across machines. */
  readonly path: string;
  /** `<c>` elements carrying a `<v>`/`<is>` value, counted from the raw XML. */
  readonly rawValueCells: number;
  /** Non-empty cells in the projected FlowDoc. */
  readonly docCells: number;
  /** Top-level body elements in the projected FlowDoc. */
  readonly bodyElements: number;
  /** Losses the reader reported. */
  readonly losses: number;
  /** Present when the read threw; the message, truncated. */
  readonly error?: string;
}

const decoder = new TextDecoder('utf-8');

// A value-bearing cell writes either <v> (stored value) or <is> (inline string).
// Namespace-prefix-agnostic: some producers write <x:v>. Self-closing cells
// (<c r="A1"/>) carry no value and are correctly not counted.
const VALUE_NODE = /<(?:[A-Za-z_][\w.-]*:)?(?:v|is)[\s>]/g;
// Worksheets normally live under xl/worksheets/, but producers put them
// elsewhere (tdf76115.xlsx uses xl/sheet1.xml) and the relationship graph is
// the only authority on which part is which. Rather than borrow our own
// resolver — which would make the reference metric depend on the code it is
// meant to check — recognise a worksheet by the one element only a worksheet
// has.
const XL_PART = /^xl\/.+\.xml$/i;
const SHEET_DATA = /<(?:[A-Za-z_][\w.-]*:)?sheetData[\s>/]/;
// §18.14 external-link parts wrap CACHED values of another workbook's cells in
// <sheetDataSet><sheetData sheetId=…>. They match the worksheet probe above but
// are not cells of this document and Excel renders none of them, so counting
// them would make a correct reader look like it dropped content.
const EXTERNAL_LINK_PART = /^xl\/externalLinks\//i;

/**
 * Count the value-bearing cells straight from the package XML — the reference
 * the projected {@link FileProbe.docCells} is measured against. Deliberately
 * independent of our parser: it must stay true even when the parser drops
 * everything.
 */
function countRawValueCells(bytes: Uint8Array): number {
  const pkg = OpcPackage.open(bytes);
  let total = 0;
  for (const part of pkg.listParts()) {
    if (!XL_PART.test(part) || EXTERNAL_LINK_PART.test(part)) continue;
    const xml = decoder.decode(pkg.requirePart(part));
    if (!SHEET_DATA.test(xml)) continue;
    total += xml.match(VALUE_NODE)?.length ?? 0;
  }
  return total;
}

/** Concatenated text of a body element subtree. */
function textOf(elements: ReadonlyArray<BodyElement>): string {
  let out = '';
  for (const element of elements) {
    if (element.kind === 'paragraph') {
      for (const run of element.paragraph.runs) out += run.text;
    } else if (element.kind === 'table') {
      for (const row of element.table.rows) {
        for (const cell of row.cells) out += textOf(cell.content);
      }
    }
  }
  return out;
}

/**
 * Count the cells that carry text in the projected document.
 *
 * Deliberately does NOT trim. The reference this is measured against counts
 * `<v>`/`<is>` nodes, which makes no judgement about whitespace, so trimming
 * here would make the two sides disagree about the same cell. Spreadsheets are
 * full of formulas returning a single space — tdf171828.xlsx has ~1550 of them,
 * and trimming reported that file as losing 55% of its content when nothing was
 * lost at all.
 */
function countDocCells(doc: FlowDoc): number {
  let total = 0;
  const walk = (elements: ReadonlyArray<BodyElement>): void => {
    for (const element of elements) {
      if (element.kind !== 'table') continue;
      for (const row of element.table.rows) {
        for (const cell of row.cells) {
          if (textOf(cell.content).length > 0) total++;
          walk(cell.content);
        }
      }
    }
  };
  walk(doc.body);
  return total;
}

function probe(path: string): FileProbe {
  const bytes = new Uint8Array(readFileSync(path));
  // Counted first and outside the try: when the read throws we still want to
  // know how much content was actually in the file.
  let rawValueCells = 0;
  try {
    rawValueCells = countRawValueCells(bytes);
  } catch {
    // A package we cannot even open has no countable cells; the read below
    // reports the real error.
  }
  try {
    const { doc, losses } = readXlsx(bytes);
    return {
      path,
      rawValueCells,
      docCells: countDocCells(doc),
      bodyElements: doc.body.length,
      losses: losses.length,
    };
  } catch (err) {
    return {
      path,
      rawValueCells,
      docCells: 0,
      bodyElements: 0,
      losses: 0,
      error: (err as Error).message.slice(0, 120),
    };
  }
}

function main(): void {
  const [out, ...files] = process.argv.slice(2);
  if (out === undefined) throw new Error('usage: xlsx-invariant-child <jsonl-out> <file>...');
  for (const file of files) {
    // Written BEFORE the parse so a fatal OOM still names its culprit.
    appendFileSync(out, JSON.stringify({ phase: 'start', path: file }) + '\n');
    appendFileSync(out, JSON.stringify({ phase: 'done', ...probe(file) }) + '\n');
  }
}

main();
