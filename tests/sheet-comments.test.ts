// E-SHEET W7 — cell comments / notes. Legacy notes (xl/comments) and modern
// threaded comments (xl/threadedComments + xl/persons) are read through the
// worksheet relationships and listed in a "Comments" section after the grid
// (Excel's "print comments at end of sheet"): a heading + one line per comment,
// each "<ref> — <author>: <text>". Render-only — not written back.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { Loss } from '@/core/ir/loss';
import type { BodyElement, ShapeBlock } from '@/core/document-model';
import { parseLegacyComments, parsePersons, parseThreadedComments } from '@/excel/comments-parser';
import { parseVmlDrawing } from '@/excel/vml-drawing';
import { Ream } from '@/core/converter/ream';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { convertXlsxToPdfSync } from '@/core/converter';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const TC = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

// The flattened text of every paragraph in the body (after the grid table).
function paragraphTexts(body: ReadonlyArray<BodyElement>): Array<string> {
  const out: Array<string> = [];
  for (const el of body) {
    if (el.kind === 'paragraph') out.push(el.paragraph.runs.map((r) => r.text).join(''));
  }
  return out;
}

describe('comment parsers (E-SHEET W7)', () => {
  it('parses legacy notes and resolves authorId, stripping the author prefix', () => {
    const xml = `<comments xmlns="${M}"><authors><author>Ada</author><author>Bo</author></authors>
      <commentList>
        <comment ref="B2" authorId="1"><text><r><t>Bo:</t></r><r><t xml:space="preserve">
check this</t></r></text></comment>
        <comment ref="C3" authorId="0"><text><r><t>looks good</t></r></text></comment>
      </commentList></comments>`;
    const out = parseLegacyComments(enc(xml));
    // The listing reads the author apart; the note's box shows the text as
    // written, run by run.
    expect(out).toEqual([
      {
        ref: 'B2',
        author: 'Bo',
        text: 'check this',
        runs: [{ text: 'Bo:' }, { text: '\ncheck this' }],
        threaded: false,
      },
      {
        ref: 'C3',
        author: 'Ada',
        text: 'looks good',
        runs: [{ text: 'looks good' }],
        threaded: false,
      },
    ]);
  });

  it("keeps each run's own font for the note's box", () => {
    const xml = `<comments xmlns="${M}"><authors><author>Ada</author></authors><commentList>
      <comment ref="A1" authorId="0"><text>
        <r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Ada:</t></r>
        <r><rPr><sz val="9"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">
plain</t></r></text></comment></commentList></comments>`;
    const [note] = parseLegacyComments(enc(xml));
    expect(note!.runs).toEqual([
      { text: 'Ada:', bold: true, italic: false, strike: false, sizePt: 9, fontName: 'Tahoma' },
      { text: '\nplain', bold: false, italic: false, strike: false, sizePt: 9, fontName: 'Tahoma' },
    ]);
  });

  it('resolves threaded comments through the person directory', () => {
    const persons = parsePersons(
      enc(
        `<personList xmlns="${TC}"><person displayName="Ada" id="p1"/><person displayName="Bo" id="p2"/></personList>`,
      ),
    );
    const out = parseThreadedComments(
      enc(
        `<ThreadedComments xmlns="${TC}"><threadedComment ref="A1" id="{1}" personId="p1"><text>first</text></threadedComment><threadedComment ref="A1" id="{2}" personId="p2"><text>reply</text></threadedComment></ThreadedComments>`,
      ),
      persons,
    );
    expect(out).toEqual([
      { ref: 'A1', author: 'Ada', text: 'first', threaded: true },
      { ref: 'A1', author: 'Bo', text: 'reply', threaded: true },
    ]);
  });
});

// A note Excel SHOWS (its VML shape carries <x:Visible/>) stands on the sheet
// as a pale box with a line to its cell's top-right corner — the box where its
// <x:Anchor> says, in cells and pixels. Measured against Excel's own PDF of a
// probe workbook printed `asDisplayed` (2026-10-01).
describe('shown notes (E-SHEET W7)', () => {
  // B2's note shown from C1 (+15px, +10px) to E5 (+15px, +16px); A1's hidden.
  // Default tracks: 48pt columns, 15pt rows.
  const book = (extra: Parameters<typeof buildXlsx>[0] = {}): Uint8Array =>
    buildXlsx({
      rows: [
        ['a', 'b'],
        ['c', 'd'],
      ],
      comments: [
        { ref: 'B2', author: 'Ada', text: 'shown', shownAnchor: '2, 15, 0, 10, 4, 15, 4, 16' },
        { ref: 'A1', author: 'Bo', text: 'hidden' },
      ],
      ...extra,
    });

  const shapesOf = (body: ReadonlyArray<BodyElement>): Array<ShapeBlock> =>
    body.flatMap((el) => (el.kind === 'shape' ? [el.shape] : []));

  it("reads a note shape's paint, the system colours older files name included", () => {
    const vml = (fill: string, visible: string): Uint8Array =>
      enc(
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
          `<v:shape id="_x0000_s1025" fillcolor="${fill}" strokecolor="#c00">` +
          '<v:shadow on="t"/><x:ClientData ObjectType="Note">' +
          `<x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor><x:Row>3</x:Row><x:Column>2</x:Column>${visible}` +
          '</x:ClientData></v:shape></xml>',
      );
    const [shown] = parseVmlDrawing(vml('infoBackground [80]', '<x:Visible/>')).notes;
    expect(shown).toEqual({
      row: 3,
      column: 2,
      visible: true,
      anchor: [1, 15, 0, 2, 3, 15, 4, 16],
      fillHex: 'FFFFE1',
      lineHex: 'CC0000',
      shadow: true,
    });
    // A note is not a form control, and a hidden one is still a note.
    const drawing = parseVmlDrawing(vml('#ffffe1', ''));
    expect(drawing.controls).toHaveLength(0);
    expect(drawing.notes[0]!.visible).toBe(false);
  });

  it("reads where a shown note's box stands, and the corner its line runs to", () => {
    const [shown, hidden] = readXlsxToSheetDoc(book()).sheets[0]!.comments!;
    expect(shown!.shown).toEqual({
      xPt: 2 * 48 + 15 * 0.75,
      yPt: 10 * 0.75,
      widthPt: 2 * 48,
      heightPt: 4 * 15 + 16 * 0.75 - 10 * 0.75,
      cornerXPt: 2 * 48,
      cornerYPt: 15,
      fillHex: 'FFFFE1',
      lineHex: '000000',
      shadow: true,
      textAlign: 'left',
    });
    expect(hidden!.shown).toBeUndefined();
  });

  it('draws it on a screen: a box with a shadow, and a line ending at the cell', () => {
    const losses: Array<Loss> = [];
    const flow = projectSheetDoc(readXlsxToSheetDoc(book()), { screen: true, losses });
    const shapes = shapesOf(flow.body);
    // The hidden note draws nothing; the shown one is a line and a box.
    expect(shapes).toHaveLength(2);
    const [line, box] = shapes;
    expect(box!.geometry.preset).toBe('rect');
    expect(box!.fill).toEqual({ kind: 'solid', colorHex: 'FFFFE1' });
    expect(box!.float?.posH?.offsetPt).toBeCloseTo(107.25);
    expect(box!.float?.posV?.offsetPt).toBeCloseTo(7.5);
    expect(box!.width).toBeCloseTo(96);
    expect(box!.shadow).toMatchObject({ dxPt: 2, dyPt: 2, blurPt: 0, colorHex: '000000' });
    expect(paragraphTexts(box!.text!.content)).toEqual(['shown']);
    // From the box's near edge to B2's top-right corner, the arrowhead there:
    // the line starts on the right, so it runs flipped.
    expect(line!.geometry.preset).toBe('line');
    expect(line!.float?.posH?.offsetPt).toBeCloseTo(96);
    expect(line!.width).toBeCloseTo(15 * 0.75);
    expect(line!.float?.posV?.offsetPt).toBeCloseTo(15);
    expect(line!.line?.tailEnd?.type).toBe('triangle');
    expect(line!.transform?.flipH).toBe(true);
    // A window shows nothing of the hidden one — and says so.
    expect(losses.map((l) => l.detail).join('\n')).toContain('1 cell note(s) not shown');
    expect(paragraphTexts(flow.body)).not.toContain('Comments');
  });

  it('prints it where it stands when the sheet prints notes as displayed, and lists none', () => {
    const flow = projectSheetDoc(
      readXlsxToSheetDoc(book({ pageSetup: { cellComments: 'asDisplayed' } })),
    );
    expect(shapesOf(flow.body).some((s) => s.fill.colorHex === 'FFFFE1')).toBe(true);
    expect(paragraphTexts(flow.body)).not.toContain('Comments');
  });

  it('lists every note at the end instead, and draws none, when the sheet asks for that', () => {
    const flow = projectSheetDoc(
      readXlsxToSheetDoc(book({ pageSetup: { cellComments: 'atEnd' } })),
    );
    expect(shapesOf(flow.body)).toHaveLength(0);
    expect(paragraphTexts(flow.body)).toContain('B2 — Ada: shown');
    expect(paragraphTexts(flow.body)).toContain('A1 — Bo: hidden');
  });

  it('mirrors it on a sheet that reads from the right, its line turned round', () => {
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [
          {
            name: 'R',
            rows: [
              ['a', 'b'],
              ['c', 'd'],
            ],
            rightToLeft: true,
          },
        ],
        comments: [
          { ref: 'B2', author: 'Ada', text: 'shown', shownAnchor: '2, 15, 0, 10, 4, 15, 4, 16' },
        ],
      }),
    );
    const flow = projectSheetDoc(doc, { screen: true });
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('no grid');
    // The grid reaches out to the note, so the mirror has it all to stand in.
    const width = table.table.grid.reduce((sum, w) => sum + w, 0);
    expect(width).toBeGreaterThanOrEqual(107.25 + 96);
    const [line, box] = shapesOf(flow.body);
    expect(box!.float?.posH?.offsetPt).toBeCloseTo(width - 107.25 - 96);
    // B2's corner is its LEFT one now, to the right of the box: the line runs
    // from the box's right edge to it, unflipped.
    expect(line!.float?.posH?.offsetPt).toBeCloseTo(width - 107.25);
    expect(line!.transform?.flipH).toBeUndefined();
  });
});

describe('cell comments — end to end (E-SHEET W7)', () => {
  it('lists legacy notes in a Comments section after the grid', () => {
    // §18.3.1.63 `cellComments="atEnd"` — Excel's "print comments at end of
    // sheet". Without it a note is an editing annotation and stays off the page.
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        comments: [
          { ref: 'A1', author: 'Ada', text: 'the first note' },
          { ref: 'B2', author: 'Bo', text: 'a second note' },
        ],
        pageSetup: { cellComments: 'atEnd' },
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('Comments');
    expect(texts).toContain('A1 — Ada: the first note');
    expect(texts).toContain('B2 — Bo: a second note');
  });

  it('prints no notes unless the sheet asks for them, and says it dropped them', () => {
    // §18.3.1.63 `cellComments` defaults to `none`, which is the print dialog's
    // own default. tdf171828.xlsx says `none` outright and NamedSheetViews.xlsx
    // says nothing at all; neither reference prints a word of either, and we
    // printed both — a whole page of notes on the first.
    for (const pageSetup of [{ cellComments: 'none' as const }, { paperSize: 9 }]) {
      const losses: Array<Loss> = [];
      const doc = readXlsxToSheetDoc(
        buildXlsx({
          rows: [['data']],
          comments: [{ ref: 'A1', author: 'Ada', text: 'the first note' }],
          pageSetup,
        }),
      );
      const flow = projectSheetDoc(doc, { losses });
      expect(paragraphTexts(flow.body)).not.toContain('Comments');
      // Not printed is not the same as unnoticed.
      expect(losses.map((l) => l.detail).join('\n')).toContain('cell note(s) not printed');
    }
  });

  it('lists threaded comments with their resolved authors', () => {
    const flow = Ream.parse(
      buildXlsx({
        pageSetup: { cellComments: 'atEnd' },
        rows: [['data']],
        threadedComments: [
          { ref: 'A1', personId: 'p1', text: 'question?' },
          { ref: 'A1', personId: 'p2', text: 'answer.' },
        ],
        persons: [
          { id: 'p1', name: 'Ada' },
          { id: 'p2', name: 'Bo' },
        ],
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('A1 — Ada: question?');
    expect(texts).toContain('A1 — Bo: answer.');
  });

  it('adds no Comments section to a sheet without comments (byte-zero)', () => {
    const flow = Ream.parse(buildXlsx({ rows: [['data']] })).flow;
    expect(paragraphTexts(flow.body)).not.toContain('Comments');
  });

  it('renders a commented sheet to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(
      buildXlsx({ rows: [['x']], comments: [{ ref: 'A1', author: 'Ada', text: 'note' }] }),
      {
        fonts: {
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
          bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
        },
      },
    );
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
