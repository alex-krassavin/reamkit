// A column set against both edges runs its every line to the measure but a
// paragraph's last, so a line short of it ended its paragraph, however nearly
// full: comments.pdf's "…break even after running a trace 270 times." stops
// twenty-seven points short, and "The other VMs we compared…", indented under
// it, came back joined to it.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import type { BodyElement } from '@/core/document-model';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { standardWidth } from '@/pdf-reader/standard-widths';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

/** A letter page of `ops` in Helvetica. */
function page(ops: ReadonlyArray<string>): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 1 });
  const pagesRef = doc.add(pagesMap);
  const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
  const leaf = doc.add(
    dict({
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, 612, 792],
      Resources: dict({ Font: dict({ F1: font }) }),
      Contents: content,
    }),
  );
  (pagesMap.get('Kids') as Array<PdfValue>).push(leaf);
  return doc.build(doc.add(dict({ Type: name('Catalog'), Pages: pagesRef })));
}

/** How wide `text` is set in 10pt Helvetica. */
const widthOf = (text: string): number =>
  [...text].reduce((sum, c) => sum + (standardWidth('Helvetica', c.charCodeAt(0), c) ?? 0), 0) *
  0.01;

/** How far a line runs at the face's own spaces. */
const natural = (words: ReadonlyArray<string>): number =>
  words.reduce((sum, w) => sum + widthOf(w), 0) + (words.length - 1) * widthOf(' ');

/** `words` filled into lines no wider than `width`, a word at a time. */
function fill(words: ReadonlyArray<string>, width: number): Array<Array<string>> {
  const lines: Array<Array<string>> = [[]];
  for (const w of words) {
    const line = lines[lines.length - 1]!;
    if (line.length > 0 && natural([...line, w]) > width) lines.push([w]);
    else line.push(w);
  }
  return lines;
}

/**
 * `lines` set from `top` at x 72 (the first `indent` further in), a word at a
 * time as TeX sets them: each but the last stretched to end at `right` where
 * `justified`, at the face's own spaces where not.
 */
function set(
  lines: ReadonlyArray<ReadonlyArray<string>>,
  top: number,
  right: number,
  justified: boolean,
  indent = 0,
  margin = 72,
): Array<string> {
  const ops: Array<string> = [];
  lines.forEach((words, k) => {
    const last = k === lines.length - 1;
    const left = margin + (k === 0 ? indent : 0);
    const spare = right - left - natural(words);
    const gap = widthOf(' ') + (justified && !last ? spare / (words.length - 1) : 0);
    let x = left;
    for (const w of words) {
      ops.push(`1 0 0 1 ${x.toFixed(2)} ${String(top - k * 12)} Tm (${w}) Tj`);
      x += widthOf(w) + gap;
    }
  });
  return ops;
}

const WORDS = (
  'A justified paragraph sets its lines against both edges of the column it is ' +
  'set in, every one of them but the last, which ends wherever its words end ' +
  'short of the measure, and the lines over it run to the edge because the ' +
  'spaces between their words were stretched to take up the white'
).split(' ');

/** The paragraphs a page of `ops` reads as, by their words. */
const paragraphsOf = (ops: ReadonlyArray<string>): Array<string> =>
  reconstructByLayout(PdfFile.parse(page(['BT /F1 10 Tf', ...ops, 'ET'])))
    .doc.body.filter(
      (el): el is Extract<BodyElement, { kind: 'paragraph' }> =>
        el.kind === 'paragraph' && el.paragraph.runs.length > 0,
    )
    .map((el) => el.paragraph.runs.map((r) => r.text).join(''));

describe('a column set against both edges', () => {
  const first = fill(WORDS, 300);
  // Its last line one that very nearly fills the measure.
  const nearlyFull = [...first.slice(0, -1), first[0]!];
  const second = fill([...WORDS].reverse(), 290);
  const top = 700 - nearlyFull.length * 12;

  it('ends a paragraph at a line short of the measure, however nearly full', () => {
    const ops = [...set(nearlyFull, 700, 72 + 305, true), ...set(second, top, 72 + 305, true, 12)];
    expect(paragraphsOf(ops)).toHaveLength(2);
  });

  it('runs the nearly full lines of a ragged column on, as it did', () => {
    // Set ragged, the same lines say nothing: a line of a ragged column ends
    // short of the measure whether it ends its paragraph or not.
    const ops = [
      ...set(nearlyFull, 700, 72 + 305, false),
      ...set(second, top, 72 + 305, false, 12),
    ];
    expect(paragraphsOf(ops)).toHaveLength(1);
  });

  it('keeps a quotation set in from both sides whole', () => {
    // freeculture.pdf sets its quotations in from both edges and justified to
    // their own measure: every line of one stops short of the column's, and
    // read as paragraph ends they came back a paragraph a line.
    const quote = fill(['Quoted', ...WORDS], 215);
    // Lines enough over it that the column shows it is set against both edges.
    const over = [...first.slice(0, -1), ...first.slice(0, -1), ...first.slice(0, -1)];
    const quoteTop = 700 - (over.length + 1) * 12;
    const ops = [
      ...set(over, 700, 72 + 305, true),
      ...set(quote, quoteTop, 72 + 245, true, 0, 92),
      ...set(second, quoteTop - quote.length * 12, 72 + 305, true, 12),
    ];
    const quoted = paragraphsOf(ops).filter((p) => p.includes('Quoted'));
    expect(quoted).toHaveLength(1);
    expect(quoted[0]).toMatch(/white$/u);
  });
});
