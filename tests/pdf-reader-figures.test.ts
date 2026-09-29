// §8.5 — a figure a page draws is one drawing with the words set on it
// (src/pdf-reader/figures.ts): found before the page's text is read, its words
// taken off the text as its labels, and set in the flow as one group where the
// page set it. comments.pdf's state machine came back as twenty-odd lines of
// labels in the column and its boxes and arrows a page later.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import type { BodyElement, ShapeBlock } from '@/core/document-model';
import type { PathSeg, TextRun } from '@/pdf-reader/content';
import type { PdfVector } from '@/pdf-reader/vector';
import { Ream } from '@/core/converter/ream';
import { OpcPackage } from '@/core/opc';
import { PdfFile } from '@/pdf-reader/document';
import { pageFigures } from '@/pdf-reader/figures';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

const SHEET = { width: 612, height: 792 };

/** A run of words on a baseline (y-up), a character half its size wide. */
function run(text: string, x: number, y: number, size = 9): TextRun {
  return {
    text,
    x,
    y,
    endX: x + text.length * size * 0.5,
    endY: y,
    fontSizePt: size,
    fontKey: 'F',
    colorHex: '000000',
  };
}

let painted = 0;

/** A path through `points`, its box taken from them; `curved` bends its first side. */
function path(points: ReadonlyArray<[number, number]>, curved = false, key?: number): PdfVector {
  const [first, ...rest] = points;
  const segs: Array<PathSeg> = [{ op: 'move', x: first![0], y: first![1] }];
  rest.forEach(([x, y], k) => {
    if (curved && k === 0) {
      segs.push({ op: 'cubic', x1: first![0], y1: y, x2: first![0], y2: y, x, y });
    } else segs.push({ op: 'line', x, y });
  });
  const xs = points.map((p) => p[0]);
  const ys = points.map((p) => p[1]);
  return {
    orderKey: [key ?? painted++],
    segs,
    strokeHex: '000000',
    lineWidth: 0.5,
    minX: Math.min(...xs),
    minY: Math.min(...ys),
    maxX: Math.max(...xs),
    maxY: Math.max(...ys),
  };
}

/** A box, lower-left at (x, y); `curved` rounds its first corner. */
const box = (x: number, y: number, w = 50, h = 20, curved = true): PdfVector =>
  path(
    [
      [x, y],
      [x + w, y],
      [x + w, y + h],
      [x, y + h],
      [x, y],
    ],
    curved,
  );

/**
 * Three boxes joined by two arrows — seven paths, drawn: a diagram's shape.
 * Lower-left at (x, y); it is 190 wide and 60 tall.
 */
function diagram(x: number, y: number): Array<PdfVector> {
  return [
    box(x, y + 40),
    path([
      [x + 50, y + 50],
      [x + 70, y + 35],
    ]),
    path([
      [x + 66, y + 33],
      [x + 72, y + 33],
      [x + 70, y + 38],
    ]),
    box(x + 70, y + 20),
    path([
      [x + 120, y + 30],
      [x + 140, y + 15],
    ]),
    path([
      [x + 136, y + 13],
      [x + 142, y + 13],
      [x + 140, y + 18],
    ]),
    box(x + 140, y),
  ];
}

/** The labels of {@link diagram}'s three boxes, in 5pt type. */
const labels = (x: number, y: number): Array<TextRun> => [
  run('Alpha', x + 10, y + 47, 5),
  run('Beta', x + 80, y + 27, 5),
  run('Gamma', x + 150, y + 7, 5),
];

/** Lines of 9pt text down a column from `top`, 11pt apart. */
const column = (x: number, top: number, count: number, prefix: string): Array<TextRun> =>
  Array.from({ length: count }, (_, k) =>
    run(`${prefix}${String(k + 1).padStart(2, '0')} words of the column`, x, top - k * 11),
  );

describe('the figures a page draws (§8.5)', () => {
  it('takes the words on a drawing as its labels, and leaves the text beside it', () => {
    const text = column(54, 600, 20, 'L');
    const figures = pageFigures(diagram(330, 400), [], [...text, ...labels(330, 400)], SHEET);
    expect(figures).toHaveLength(1);
    expect(figures[0]!.labels.map((r) => r.text)).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(figures[0]!.vectors).toHaveLength(7);
  });

  it('takes the spaces between a label’s words with it', () => {
    const space = run(' ', 345, 447, 5);
    const figures = pageFigures(diagram(330, 400), [], [...labels(330, 400), space], SHEET);
    expect(figures[0]!.labels).toContain(space);
  });

  it('finds none in a ruling: boxes and straight lines are a table’s', () => {
    const cells = Array.from({ length: 8 }, (_, k) =>
      box(100 + (k % 4) * 50, 500 + Math.floor(k / 4) * 20, 50, 20, false),
    );
    const words = cells.map((c, k) => run(`cell${String(k)}`, c.minX + 5, c.minY + 6));
    expect(pageFigures(cells, [], words, SHEET)).toEqual([]);
  });

  it('finds none where the words run across the drawing: that is prose on a ground', () => {
    const prose = [0, 1, 2].map((k) =>
      run('a line of prose that runs across the whole of the panel', 330, 448 - k * 11, 7),
    );
    expect(pageFigures(diagram(330, 400), [], prose, SHEET)).toEqual([]);
  });

  it('finds none where the words are set larger than the text', () => {
    // bug1771477.pdf is a web page printed: its cards are drawn round 17pt words.
    const text = column(54, 600, 20, 'L');
    const big = [run('hello world', 340, 447, 17)];
    expect(pageFigures(diagram(330, 400), [], [...text, ...big], SHEET)).toEqual([]);
  });

  it('makes one figure of parts set apart with nothing between, and two of parts a caption parts', () => {
    const text = column(54, 700, 30, 'L');
    const side = pageFigures(
      [...diagram(40, 400), ...diagram(250, 400)],
      [],
      [...text, ...labels(40, 400), ...labels(250, 400)],
      SHEET,
    );
    expect(side).toHaveLength(1);
    const caption = run('Figure 1. The first of the two diagrams', 40, 480);
    const stacked = pageFigures(
      [...diagram(300, 500), ...diagram(300, 400)],
      [],
      [...text, ...labels(300, 500), ...labels(300, 400), caption],
      SHEET,
    );
    expect(stacked).toHaveLength(2);
  });

  it('takes a few paths standing near a figure as its own', () => {
    // comments.pdf's Figure 8 draws its outer tree's head five points off the
    // rest, too few paths to be a figure of its own.
    const head = [box(525, 470, 10, 10), box(527, 482, 6, 6)];
    const figures = pageFigures([...diagram(330, 400), ...head], [], labels(330, 400), SHEET);
    expect(figures).toHaveLength(1);
    expect(figures[0]!.vectors).toEqual(expect.arrayContaining(head));
  });

  it('leaves out what an annotation draws over the page', () => {
    const bar = path(
      [
        [20, 455],
        [600, 455],
        [600, 470],
        [20, 470],
      ],
      false,
      Number.MAX_SAFE_INTEGER,
    );
    const figures = pageFigures([...diagram(330, 400), bar], [], labels(330, 400), SHEET);
    expect(figures).toHaveLength(1);
    expect(figures[0]!.vectors).not.toContain(bar);
    expect(figures[0]!.minX).toBeGreaterThan(300);
  });
});

/** A page of Courier columns (every line fills its column) with a diagram in the right one. */
function paperPdf(): Uint8Array {
  const doc = new PdfDocument();
  const courier = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Courier') }),
  );
  const helvetica = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  // 44 characters of 9pt Courier fill 237.6 points; a paragraph's first line
  // is set in two characters and ends where the others do — justified.
  const line = (tag: string, first: boolean): string => {
    const words = `${tag} ${'set in the column and filling it '.repeat(3)}`;
    return words.slice(0, first ? 42 : 44);
  };
  const ops: Array<string> = ['BT /F1 9 Tf'];
  const set = (x: number, y: number, tag: string, k: number): void => {
    const first = k % 5 === 0;
    ops.push(`1 0 0 1 ${String(x + (first ? 10.8 : 0))} ${String(y)} Tm (${line(tag, first)}) Tj`);
  };
  for (let k = 0; k < 40; k++) set(54, 720 - k * 11, `L${String(k + 1).padStart(2, '0')}`, k);
  for (let k = 0; k < 15; k++) set(317, 720 - k * 11, `R${String(k + 1).padStart(2, '0')}`, k);
  for (let k = 0; k < 15; k++) set(317, 450 - k * 11, `S${String(k + 1).padStart(2, '0')}`, k);
  ops.push('ET');
  // The diagram: three rounded boxes, two arrows, a word in each box.
  const round = (x: number, y: number): string =>
    `${String(x + 3)} ${String(y)} m ${String(x + 47)} ${String(y)} l ` +
    `${String(x + 50)} ${String(y)} ${String(x + 50)} ${String(y)} ${String(x + 50)} ${String(y + 3)} c ` +
    `${String(x + 50)} ${String(y + 20)} l ${String(x)} ${String(y + 20)} l ${String(x)} ${String(y + 3)} l ` +
    `${String(x)} ${String(y)} ${String(x)} ${String(y)} ${String(x + 3)} ${String(y)} c h S`;
  const arrow = (x: number, y: number): string =>
    `${String(x)} ${String(y)} m ${String(x + 18)} ${String(y - 13)} l S ` +
    `${String(x + 15)} ${String(y - 17)} m ${String(x + 21)} ${String(y - 17)} l ${String(x + 19)} ${String(y - 11)} l h f`;
  ops.push(
    '0.5 w',
    round(330, 520),
    arrow(380, 530),
    round(400, 500),
    arrow(450, 510),
    round(470, 480),
  );
  ops.push('BT /F2 5 Tf 1 0 0 1 340 527 Tm (Alpha) Tj 1 0 0 1 410 507 Tm (Beta) Tj');
  ops.push('1 0 0 1 480 487 Tm (Gamma) Tj ET');
  const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 1 });
  const pagesRef = doc.add(pagesMap);
  const page = doc.add(
    dict({
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, 612, 792],
      Resources: dict({ Font: dict({ F1: courier, F2: helvetica }) }),
      Contents: content,
    }),
  );
  (pagesMap.get('Kids') as Array<PdfValue>).push(page);
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

const textOf = (el: BodyElement): string =>
  el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '';

const figureIn = (body: ReadonlyArray<BodyElement>): ShapeBlock | undefined =>
  body.flatMap((el) => (el.kind === 'shape' && el.shape.children ? [el.shape] : []))[0];

const words = (shape: ShapeBlock): Array<string> =>
  (shape.children ?? []).flatMap((c) =>
    (c.shape.text?.content ?? []).map((el) => textOf(el)).filter((t) => t !== ''),
  );

describe('a figure in the flow (§20.5.2.17)', () => {
  const { doc } = reconstructByLayout(PdfFile.parse(paperPdf()));
  const body = doc.body;

  it('sets the drawing and its labels as one group, not the labels as lines of text', () => {
    expect(body.map(textOf).join(' ')).not.toMatch(/Alpha|Beta|Gamma/u);
    const figure = figureIn(body);
    expect(figure).toBeDefined();
    expect(figure!.float).toBeUndefined();
    expect(words(figure!)).toEqual(['Alpha', 'Beta', 'Gamma']);
    // …each label in a box that does not wrap: set in a wider face, a word
    // would break in two.
    const boxes = figure!.children!.filter((c) => c.shape.text !== undefined);
    expect(boxes.every((c) => c.shape.text!.noWrap === true)).toBe(true);
  });

  it('stands between the text over it and the text under it, in its column', () => {
    const at = body.findIndex((el) => el.kind === 'shape' && el.shape.children !== undefined);
    const over = body.findIndex((el) => textOf(el).includes('R15'));
    const under = body.findIndex((el) => textOf(el).includes('S01'));
    expect(over).toBeGreaterThan(body.findIndex((el) => textOf(el).includes('L40')));
    expect(over).toBeLessThan(at);
    expect(under).toBe(at + 1);
    // The text under it is spaced from its foot, not from the text over it:
    // counted from there, the figure's height would be taken twice.
    const next = body[under]!;
    const before = next.kind === 'paragraph' ? (next.paragraph.properties.spacingBefore ?? 0) : 0;
    expect(before).toBeCloseTo(480 - (450 + 0.8 * 11), 1);
    expect(figureIn(body)!.height).toBeGreaterThan(55);
  });

  it('reads a justified page in its columns', () => {
    // Every full line of a justified column ends where the column does, so
    // its right edges agree more than its left ones — which is what a column
    // of amounts looks like, and the page was read straight across.
    expect(doc.sections.some((s) => s.properties.columns?.count === 2)).toBe(true);
  });

  it('is written as a group of shapes and pictures, and read back as one', async () => {
    const docx = await Ream.parse(paperPdf()).convert('docx');
    const xml = new TextDecoder().decode(OpcPackage.open(docx).getMainDocument().data);
    expect(xml).toContain('<wpg:wgp');
    expect(xml).toMatch(/<a:chOff x="0" y="0"\/><a:chExt cx="\d+" cy="\d+"\/>/u);
    expect(xml).toContain('wrap="none"');
    const again = figureIn(Ream.parse(docx).flow.body);
    expect(again?.children?.length).toBe(figureIn(body)!.children!.length);
    expect(words(again!)).toEqual(['Alpha', 'Beta', 'Gamma']);
  });
});
