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

import type { BodyElement, FlowDoc } from '@/core/document-model';
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
const WORKSHEET_PART = /^xl\/worksheets\/[^/]+\.xml$/i;

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
    if (!WORKSHEET_PART.test(part)) continue;
    const xml = decoder.decode(pkg.requirePart(part));
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

/** Count the cells that carry visible text in the projected document. */
function countDocCells(doc: FlowDoc): number {
  let total = 0;
  const walk = (elements: ReadonlyArray<BodyElement>): void => {
    for (const element of elements) {
      if (element.kind !== 'table') continue;
      for (const row of element.table.rows) {
        for (const cell of row.cells) {
          if (textOf(cell.content).trim().length > 0) total++;
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
