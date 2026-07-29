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
// (<c r="A1"/>) carry no value and are correctly not counted. NOT global:
// `.test()` on a /g regex carries lastIndex between calls and would skip every
// other cell.
const HAS_VALUE = /<(?:[A-Za-z_][\w.-]*:)?(?:v|is)[\s>]/;
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
// §18.2.19 `<sheet state="hidden">`. A hidden tab is content the DOCUMENT says
// not to print — Excel and LibreOffice both leave it out — so counting its
// cells marks a reader that correctly omits it as having lost something. The
// relationship id is the only link from the tab to its part, so resolve through
// the workbook's own rels rather than guessing at file names.
const SHEET_ELEMENT = /<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*\/?>/g;
const RELATIONSHIP_ELEMENT = /<Relationship\b[^>]*\/?>/g;

/**
 * Count the value-bearing cells straight from the package XML — the reference
 * the projected {@link FileProbe.docCells} is measured against. Deliberately
 * independent of our parser: it must stay true even when the parser drops
 * everything.
 */
function countRawValueCells(bytes: Uint8Array): number {
  const pkg = OpcPackage.open(bytes);
  const hiddenParts = hiddenSheetParts(pkg);
  let total = 0;
  for (const part of pkg.listParts()) {
    if (!XL_PART.test(part) || EXTERNAL_LINK_PART.test(part)) continue;
    if (hiddenParts.has(part)) continue;
    const xml = decoder.decode(pkg.requirePart(part));
    if (!SHEET_DATA.test(xml)) continue;
    total += countPrintableValues(xml);
  }
  return total;
}

/** The worksheet parts belonging to tabs the workbook marks hidden. */
function hiddenSheetParts(pkg: OpcPackage): Set<string> {
  const out = new Set<string>();
  const wb = pkg.getPart('xl/workbook.xml');
  const rels = pkg.getPart('xl/_rels/workbook.xml.rels');
  if (!wb || !rels) return out;
  const targets = new Map<string, string>();
  for (const tag of decoder.decode(rels).match(RELATIONSHIP_ELEMENT) ?? []) {
    const id = attr(tag, 'Id');
    const target = attr(tag, 'Target');
    if (id && target) targets.set(id, target.replace(/^\/?(xl\/)?/, 'xl/'));
  }
  for (const tag of decoder.decode(wb).match(SHEET_ELEMENT) ?? []) {
    const state = attr(tag, 'state');
    if (state !== 'hidden' && state !== 'veryHidden') continue;
    const id = attr(tag, 'r:id') ?? attr(tag, 'id');
    const target = id ? targets.get(id) : undefined;
    if (target) out.add(target);
  }
  return out;
}

// §18.3.1.13 / §18.3.1.73 `hidden`. A hidden row or column is content the
// DOCUMENT says not to print, so counting it here would mark a reader that
// correctly omits it as having lost something — the question this metric asks
// is whether we dropped anything silently, not whether we drew everything the
// file contains.
const COL_ELEMENT = /<(?:[A-Za-z_][\w.-]*:)?col\b[^>]*>/g;
const ROW_ELEMENT = /<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*>/g;
const CELL_ELEMENT = /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/>|>)/g;
const attr = (tag: string, name: string): string | undefined =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
const isHidden = (tag: string): boolean => {
  const v = attr(tag, 'hidden');
  return v === '1' || v === 'true';
};

/** Column index (0-based) of an A1 reference like `"AB12"`; -1 when unusable. */
function columnOf(ref: string | undefined): number {
  const letters = /^([A-Za-z]+)/.exec(ref ?? '')?.[1];
  if (!letters) return -1;
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Value-bearing cells that the sheet actually prints. */
function countPrintableValues(xml: string): number {
  const hiddenCols = new Set<number>();
  for (const tag of xml.match(COL_ELEMENT) ?? []) {
    if (!isHidden(tag)) continue;
    const min = Number(attr(tag, 'min') ?? NaN);
    const max = Number(attr(tag, 'max') ?? NaN);
    if (!Number.isFinite(min) || !Number.isFinite(max)) continue;
    // A hidden run to the sheet's last column is common and harmless to bound.
    for (let c = min; c <= Math.min(max, 16_384); c++) hiddenCols.add(c - 1);
  }

  // Rows are delimited by their opening tags; the slice up to the next one is
  // that row's cells.
  let total = 0;
  const rowTags = [...xml.matchAll(ROW_ELEMENT)];
  for (let i = 0; i < rowTags.length; i++) {
    const tag = rowTags[i]!;
    if (isHidden(tag[0])) continue;
    const from = tag.index + tag[0].length;
    const to = i + 1 < rowTags.length ? rowTags[i + 1]!.index : xml.length;
    for (const cell of splitCells(xml.slice(from, to))) {
      if (hiddenCols.has(columnOf(attr(cell, 'r')))) continue;
      if (HAS_VALUE.test(cell)) total++;
    }
  }
  return total;
}

/** One string per `<c>…</c>` (or self-closing `<c/>`) in a row's markup. */
function splitCells(rowBody: string): Array<string> {
  const out: Array<string> = [];
  const starts = [...rowBody.matchAll(CELL_ELEMENT)];
  for (let i = 0; i < starts.length; i++) {
    const from = starts[i]!.index;
    const to = i + 1 < starts.length ? starts[i + 1]!.index : rowBody.length;
    out.push(rowBody.slice(from, to));
  }
  return out;
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
