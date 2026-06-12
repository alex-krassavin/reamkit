// docx-writer round-trip gate (E-DOCX D6): for every docx in the corpus,
//   readDocx(bytes) → flow0 → writeDocx(flow0) → readDocx(written) → flow1
// and compare the FlowDoc signatures. The writer is denormalized but
// semantically equivalent, so the comparison is on extracted features, not
// bytes: the body/table/header-footer text must match exactly, and the block
// counts (paragraphs, tables, images, hyperlinks, bookmarks) must agree.
//
// Pure JS parse + write of our own bytes — no LibreOffice, no rendering, so no
// Docker sandbox is needed (readDocx is already zip-bomb hardened for the
// untrusted corpus inputs). Reports a markdown table to stdout.
//
// Usage: CORPUS_DIR=corpus/external/poi-docx npx tsx scripts/corpus/roundtrip.ts

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { basename, dirname, resolve } from 'node:path';

import { listCorpus } from './lib';
import type { BodyElement, FlowDoc } from '@/core/document-model';
import { readDocx } from '@/word/docx-reader';
import { writeDocx } from '@/word/docx-writer';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const corpusDir = process.env.CORPUS_DIR
  ? resolve(root, process.env.CORPUS_DIR)
  : resolve(root, 'corpus/external/poi-docx');

interface Signature {
  readonly text: string;
  readonly paragraphs: number;
  readonly tables: number;
  readonly images: number;
  readonly hyperlinks: number;
  readonly bookmarks: number;
}

function walkBlocks(els: ReadonlyArray<BodyElement>, sig: Mutable<Signature>): void {
  for (const el of els) {
    if (el.kind === 'paragraph') {
      sig.paragraphs++;
      for (const r of el.paragraph.runs) {
        if (!r.listMarker && r.inlineImage === undefined) sig.text += r.text;
        if (r.inlineImage !== undefined) sig.images++;
        if (r.href !== undefined || r.anchor !== undefined) sig.hyperlinks++;
      }
      sig.bookmarks += el.paragraph.bookmarks?.length ?? 0;
    } else if (el.kind === 'table') {
      sig.tables++;
      for (const row of el.table.rows) for (const cell of row.cells) walkBlocks(cell.content, sig);
    } else if (el.kind === 'image') {
      sig.images++;
    }
  }
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

function signature(flow: FlowDoc): Signature {
  const sig: Mutable<Signature> = {
    text: '',
    paragraphs: 0,
    tables: 0,
    images: 0,
    hyperlinks: 0,
    bookmarks: 0,
  };
  walkBlocks(flow.body, sig);
  // Headers/footers carry text too — the writer round-trips them through parts.
  // They live in a Map keyed by relationship id, and the writer reassigns those
  // ids, so iterate in a deterministic (sorted-text) order to compare the SET
  // of header/footer content rather than the Map's insertion order.
  const hfTexts: Array<string> = [];
  for (const content of flow.headersFooters?.values() ?? []) {
    const hf: Mutable<Signature> = {
      text: '',
      paragraphs: 0,
      tables: 0,
      images: 0,
      hyperlinks: 0,
      bookmarks: 0,
    };
    walkBlocks(content, hf);
    sig.paragraphs += hf.paragraphs;
    sig.tables += hf.tables;
    sig.images += hf.images;
    sig.hyperlinks += hf.hyperlinks;
    sig.bookmarks += hf.bookmarks;
    hfTexts.push(hf.text);
  }
  hfTexts.sort();
  sig.text = (sig.text + hfTexts.join('')).replace(/\s+/g, '');
  return sig;
}

interface Row {
  // ✅ identical · ⚠️ divergent · ❌ writer threw · ⏭ input unreadable (reader)
  readonly name: string;
  readonly status: '✅' | '⚠️' | '❌' | '⏭';
  readonly note: string;
}

function diffNote(a: Signature, b: Signature): string {
  const parts: Array<string> = [];
  if (a.text !== b.text) {
    const lost = a.text.length - b.text.length;
    parts.push(
      `text ${a.text.length}→${b.text.length} (${lost >= 0 ? '-' : '+'}${Math.abs(lost)})`,
    );
  }
  for (const k of ['paragraphs', 'tables', 'images', 'hyperlinks', 'bookmarks'] as const) {
    if (a[k] !== b[k]) parts.push(`${k} ${a[k]}→${b[k]}`);
  }
  return parts.join(', ');
}

function main(): void {
  const inputs = listCorpus(corpusDir);
  if (inputs.length === 0) {
    console.error(`No docx in ${corpusDir}`);
    process.exit(1);
  }
  const rows: Array<Row> = [];
  let ok = 0;
  let warn = 0;
  let fail = 0;
  let skip = 0;

  for (const input of inputs) {
    const name = basename(input);
    if (!name.endsWith('.docx')) continue;
    // The reader's own rejections (encrypted, invalid zip, …) are NOT writer
    // failures — the writer never sees those documents. Separate the phases.
    let flow0: FlowDoc;
    try {
      flow0 = readDocx(new Uint8Array(readFileSync(input))).doc;
    } catch {
      rows.push({ name, status: '⏭', note: 'input unreadable (reader)' });
      skip++;
      continue;
    }
    try {
      const written = writeDocx(flow0).bytes;
      const flow1 = readDocx(written).doc;
      const a = signature(flow0);
      const b = signature(flow1);
      const note = diffNote(a, b);
      if (note === '') {
        rows.push({ name, status: '✅', note: `${a.paragraphs}p ${a.tables}t ${a.images}i` });
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
  lines.push(`# docx-writer round-trip — ${basename(corpusDir)}`);
  lines.push('');
  lines.push(
    `**${readable} readable docs — ✅ ${ok} identical · ⚠️ ${warn} divergent · ` +
      `❌ ${fail} writer-failed; ⏭ ${skip} unreadable input.**`,
  );
  lines.push('');
  lines.push('| Doc | St | Note |');
  lines.push('|---|---|---|');
  for (const r of rows) lines.push(`| ${r.name} | ${r.status} | ${r.note} |`);
  console.log(lines.join('\n'));
}

main();
