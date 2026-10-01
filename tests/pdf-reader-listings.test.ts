// A code listing is read as the listing it is: every line of it set in a
// typewriter face (§9.8.2), ended where the page ends it, its comments lined up
// by the tab the page's white stands for — and none of that white taken for
// the gutter of a page in columns. comments.pdf's third page came back with its
// code in one column and its comments in another, and two pages long.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import type { BodyElement } from '@/core/document-model';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { extractPageText } from '@/pdf-reader/text';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

/**
 * Letter pages of `ops`, one page each, with `/F1` the given font dictionary
 * and `/F2` Helvetica.
 */
function pages(
  sheets: ReadonlyArray<ReadonlyArray<string>>,
  f1: Record<string, PdfValue>,
): Uint8Array {
  const doc = new PdfDocument();
  const mono = doc.add(dict(f1));
  const helvetica = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: sheets.length });
  const pagesRef = doc.add(pagesMap);
  for (const ops of sheets) {
    const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
    const leaf = doc.add(
      dict({
        Type: name('Page'),
        Parent: pagesRef,
        MediaBox: [0, 0, 612, 792],
        Resources: dict({ Font: dict({ F1: mono, F2: helvetica }) }),
        Contents: content,
      }),
    );
    (pagesMap.get('Kids') as Array<PdfValue>).push(leaf);
  }
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

/** A letter page of `ops`. */
const page = (ops: ReadonlyArray<string>, f1: Record<string, PdfValue>): Uint8Array =>
  pages([ops], f1);

const COURIER = { Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Courier') };

/** A simple font whose /Widths are `widths` from code 32 on, and no program. */
const widthsFont = (base: string, widths: ReadonlyArray<number>): Record<string, PdfValue> => ({
  Type: name('Font'),
  Subtype: name('Type1'),
  BaseFont: name(base),
  FirstChar: 32,
  LastChar: 31 + widths.length,
  Widths: [...widths],
});

const fixedPitchOf = (f1: Record<string, PdfValue>): boolean | undefined => {
  const file = PdfFile.parse(page(['BT /F1 10 Tf 72 700 Td (abc) Tj ET'], f1));
  return extractPageText(file, file.pages()[0]!)[0]?.fixedPitch;
};

describe('a typewriter face (§9.8.2)', () => {
  it('is one whose widths are all the same, whatever its flags say', () => {
    // pdfTeX sets no FixedPitch flag on Computer Modern's typewriter face.
    expect(
      fixedPitchOf(
        widthsFont(
          'CMTT9',
          Array.from({ length: 95 }, () => 525),
        ),
      ),
    ).toBe(true);
  });

  it('is one whose descriptor says so, whatever its widths', () => {
    const flagged = {
      ...widthsFont('Mono', [250, 333, 408, 500, 500, 833, 778, 180, 333, 333]),
      FontDescriptor: dict({ Type: name('FontDescriptor'), FontName: name('Mono'), Flags: 1 }),
    };
    expect(fixedPitchOf(flagged)).toBe(true);
  });

  it('is the standard Courier', () => {
    expect(fixedPitchOf(COURIER)).toBe(true);
  });

  it('is not a face that shows a few digits, all as wide as each other', () => {
    expect(fixedPitchOf(widthsFont('CMR6', [611, 611, 611, 611]))).toBeUndefined();
    expect(
      fixedPitchOf(widthsFont('Times', [250, 333, 408, 500, 500, 833, 778, 180, 333, 333])),
    ).toBeUndefined();
  });
});

const CELL = 5.4; // 9pt Courier

/** The listing: code at the left, comments lined up at the 24th cell. */
const LISTING: ReadonlyArray<[string, string]> = [
  ['v0 := ld state[748]', '// load primes from the trace activation record'],
  ['   st sp[0], v0', '// store primes to interpreter stack'],
  ['v1 := ld state[764]', '// load k from the trace activation record'],
  ['v2 := i2f(v1)', '// convert k from int to double'],
  ['   st sp[8], v1', '// store k'],
  ['...', ''],
  ['exit:', ''],
  ['v3 := ld v0[4]', '// load class word for primes'],
  ['v4 := and v3, -4', '// mask out object class tag for primes'],
  ['v5 := eq v4, Array', '// test whether primes is an array'],
  ['   xf v5', '// side exit if v5 is false'],
  ['v6 := call(v0)', '// call function to set array element'],
];

/** One line of code set word by word, the way TeX sets it, at `y`, from `left`. */
function codeLine(code: string, comment: string, y: number, left = 54): Array<string> {
  const ops: Array<string> = [];
  const words = (text: string, from: number): void => {
    let at = 0;
    for (const word of text.split(' ')) {
      // §7.3.4.2 — a parenthesis in a literal string is escaped.
      const shown = word.replace(/[()\\]/gu, (c) => `\\${c}`);
      if (word !== '')
        ops.push(`1 0 0 1 ${String(from + at * CELL)} ${String(y)} Tm (${shown}) Tj`);
      at += word.length + 1;
    }
  };
  words(code, left);
  if (comment !== '') words(comment, left + 24 * CELL);
  return ops;
}

/** Prose in two columns under the listing: `L..` at the left, `R..` at the right. */
function columns(top: number, rows: number): Array<string> {
  const ops: Array<string> = [];
  for (let k = 0; k < rows; k++) {
    const n = String(k + 1).padStart(2, '0');
    const y = String(top - k * 12);
    ops.push(`1 0 0 1 54 ${y} Tm (L${n} a line of prose that runs across its column) Tj`);
    ops.push(`1 0 0 1 317 ${y} Tm (R${n} a line of prose that runs across its column) Tj`);
  }
  return ops;
}

const textOf = (el: BodyElement): string =>
  el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '';

describe('a code listing across the head of a page in columns', () => {
  // Three times over: thirty-six lines of it split at one x, where the
  // columns under it are sixteen lines deep — comments.pdf's proportion.
  // …and a tail of short lines, none reaching across the page: comments.pdf's
  // second listing ends "...", "side_exit_1:" and three short lines more.
  const TAIL: ReadonlyArray<[string, string]> = [
    ['.....', ''],
    ['done:', ''],
    ['   jmp epilog', '// return'],
  ];
  const listing = [...LISTING, ...LISTING, ...LISTING, ...TAIL];
  const ops = ['BT /F1 9 Tf'];
  listing.forEach(([code, comment], k) => ops.push(...codeLine(code, comment, 760 - k * 10)));
  ops.push('/F2 9 Tf', ...columns(355, 16), 'ET');
  // The paper goes on in its own face: the listing is a part of it.
  const more = [
    'BT /F2 9 Tf',
    ...columns(740, 55).map((op) => op.replace(/\(([LR])/u, '($1x')),
    'ET',
  ];
  const { doc } = reconstructByLayout(PdfFile.parse(pages([ops, more], COURIER)));
  const texts = doc.body.map(textOf);

  it('sets each line of the code as a paragraph of its own, its comment on a stop', () => {
    const lines = [...LISTING, ...TAIL].map(([code, comment]) =>
      texts.indexOf(`${code.trim()}${comment ? `\t${comment}` : ''}`),
    );
    expect(lines.every((at) => at >= 0)).toBe(true);
    // …in the order the page sets them, the short lines among the rest.
    expect([...lines].sort((a, b) => a - b)).toEqual(lines);
    const first = doc.body[lines[0]!]!;
    expect(
      first.kind === 'paragraph' ? first.paragraph.properties.tabs?.[0]?.positionPt : 0,
    ).toBeCloseTo(24 * CELL, 1);
  });

  it('keeps the listing across the page, its short last lines too', () => {
    // Left of the gutter, they read as the first lines of the left column, and
    // came back set in the columns with the prose.
    const tail = TAIL.map(([code, comment]) =>
      texts.indexOf(`${code.trim()}${comment ? `\t${comment}` : ''}`),
    );
    const sectionOf = (at: number) => doc.sections.find((s) => s.endIndex >= at);
    expect(tail.every((at) => at >= 0 && sectionOf(at)?.properties.columns === undefined)).toBe(
      true,
    );
  });

  it('reads the prose under it in the columns it is set in', () => {
    const prose = texts.join(' ').match(/[LR]\d\d/gu) ?? [];
    const column = (letter: string): Array<string> =>
      Array.from({ length: 16 }, (_, k) => `${letter}${String(k + 1).padStart(2, '0')}`);
    expect(prose).toEqual([...column('L'), ...column('R')]);
    expect(doc.sections.some((s) => s.properties.columns?.count === 2)).toBe(true);
  });
});

describe('a code listing inside a column', () => {
  it('keeps its lines apart, however well the next would fit', () => {
    // comments.pdf's Figure 1 set line 5 under line 4, and read as prose line
    // 4 took line 5's number onto its end.
    const code = ['1 for (i = 0; i < 9; i++) {', '2   a[i] = i;', '3 }'];
    const ops = ['BT /F2 9 Tf'];
    for (let k = 0; k < 24; k++) {
      const y = String(700 - k * 12);
      ops.push(`1 0 0 1 54 ${y} Tm (a line of prose that runs across the left column) Tj`);
      // The right column: prose, then the listing, then prose again.
      if (k < 8 || k >= 12) {
        ops.push(`1 0 0 1 317 ${y} Tm (a line of prose that runs across the right column) Tj`);
      }
    }
    ops.push('/F1 9 Tf');
    code.forEach((line, k) => ops.push(...codeLine(line, '', 700 - 8 * 12 - k * 12, 317)));
    ops.push('ET');
    const { doc } = reconstructByLayout(PdfFile.parse(page(ops, COURIER)));
    const texts = doc.body.map(textOf);
    expect(texts).toContain('1 for (i = 0; i < 9; i++) {');
    expect(texts).toContain('2\ta[i] = i;');
    expect(texts).toContain('3 }');
  });
});

describe('a document typed throughout', () => {
  it('is prose in its typewriter face: its lines run on as a paragraph', () => {
    // A letter typed in Courier, a screenplay: the face is the document's own,
    // and no line of it is a listing's.
    const words = 'the lines of a letter typed in one face run on from one to the next';
    const ops = ['BT /F1 10 Tf'];
    for (let k = 0; k < 6; k++) ops.push(`1 0 0 1 72 ${String(700 - k * 12)} Tm (${words}) Tj`);
    ops.push('ET');
    const { doc } = reconstructByLayout(PdfFile.parse(page(ops, COURIER)));
    const texts = doc.body.map(textOf).filter((t) => t !== '');
    expect(texts).toHaveLength(1);
  });
});
