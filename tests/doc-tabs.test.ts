// Tabs in a legacy `.doc` ([MS-DOC] §2.6.2): a 0x09 in the text is a tab, and a
// paragraph's custom tab stops are its style chain's, as the sprmPChgTabsPapx /
// sprmPChgTabs of each style and of its own PAPX delete and add them. Read as a
// space, with no stops, the text after a tab fell wherever the space ended.
//
// Placement measured in Word for Mac on a probe of our own: a Courier New 10pt
// .docx of these stops, which LibreOffice wrote as a .doc (its sprmPChgTabsPapx
// deleting the style's stop included), opened in Word and read back through
// AppleScript (x from the text boundary, ±0.4pt). The tests check the same
// geometry in Roboto: where the text after each tab starts, centres or ends.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc } from './fixtures/build-doc';
import type { DocParaFormat, DocStyleDef } from './fixtures/build-doc';
import type { TabStop } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { TextLineItem } from '@/layout/page-doc';

import { readDoc } from '@/word/doc/doc-reader';
import { extractDocContent } from '@/word/doc/doc-text';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};
const MARGIN = 72; // the fixture's section states no margins: Word's inch

// A .doc of one paragraph per line of `text`, each with the PAPX at its index.
function docOf(
  lines: ReadonlyArray<string>,
  paras: ReadonlyArray<Omit<DocParaFormat, 'length'>>,
  styles?: ReadonlyArray<DocStyleDef>,
): FlowDoc {
  const text = lines.map((l) => `${l}\r`).join('');
  return readDoc(
    buildDoc([{ text, compressed: false }], {
      paraRuns: lines.map((l, i) => ({ length: l.length + 1, ...paras[i] })),
      ...(styles ? { styles } : {}),
    }),
  ).doc;
}

function stopsOf(doc: FlowDoc): Array<ReadonlyArray<TabStop> | undefined> {
  return doc.body.flatMap((el) => (el.kind === 'paragraph' ? [el.paragraph.properties.tabs] : []));
}

// Each laid-out line as its text tokens with where each starts and ends, in
// points from the left margin.
function placed(doc: FlowDoc): Array<Array<{ text: string; x: number; end: number }>> {
  const laid = layoutStyledDocument(doc.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(doc),
  });
  return laid.pages.flatMap((page) =>
    page.commands
      .filter((c): c is TextLineItem => c.type === 'line')
      .map((l) => {
        let x = l.originX - MARGIN;
        return l.line.tokens.flatMap((t) => {
          const at = x;
          x += t.widthPt;
          return t.kind === 'text' && !t.tab ? [{ text: t.text, x: at, end: x }] : [];
        });
      }),
  );
}

// The token holding `text` on the line it is on.
function token(doc: FlowDoc, text: string): { x: number; end: number } {
  const hit = placed(doc)
    .flat()
    .find((t) => t.text === text);
  if (!hit) throw new Error(`no token "${text}"`);
  return hit;
}

describe('doc reader: tabs', () => {
  it('keeps a tab a tab', () => {
    const doc = docOf(['ab\tL'], [{}]);
    const p = doc.body[0]!;
    expect(p.kind === 'paragraph' && p.paragraph.runs.map((r) => r.text).join('')).toBe('ab\tL');
  });

  it('reads the stops a paragraph states: where, how text sits and the leader', () => {
    // TBD: jc in the low three bits, tlc in the next three; tlcHeavy is drawn
    // as the underscore, and tlcDefault (7) is no leader.
    const doc = docOf(
      ['x'],
      [
        {
          tabs: {
            add: [
              { pos: 1440, jc: 0, tlc: 1 },
              { pos: 2880, jc: 1, tlc: 2 },
              { pos: 4320, jc: 2, tlc: 4 },
              { pos: 5040, jc: 3, tlc: 5 },
              { pos: 5760, jc: 0, tlc: 7 },
            ],
          },
        },
      ],
    );
    expect(stopsOf(doc)).toEqual([
      [
        { positionPt: 72, alignment: 'left', leader: 'dot' },
        { positionPt: 144, alignment: 'center', leader: 'hyphen' },
        { positionPt: 216, alignment: 'right', leader: 'underscore' },
        { positionPt: 252, alignment: 'decimal', leader: 'middleDot' },
        { positionPt: 288, alignment: 'left' },
      ],
    ]);
  });

  it('sets the text after a tab where Word does against each kind of stop', () => {
    // Word: L at 72.0; CCC from 135.3 to 152.7, centred on 144; RRR ending at
    // 216.0; "12.5" with its point on 144 ("1" at 132.15 in Courier) — where
    // "12" with no point ends.
    const doc = docOf(
      ['ab\tL', 'ab\tCCC', 'ab\tRRR', 'ab\t12.5', 'ab\t12'],
      [
        { tabs: { add: [{ pos: 1440 }] } },
        { tabs: { add: [{ pos: 2880, jc: 1 }] } },
        { tabs: { add: [{ pos: 4320, jc: 2 }] } },
        { tabs: { add: [{ pos: 2880, jc: 3 }] } },
        { tabs: { add: [{ pos: 2880, jc: 3 }] } },
      ],
    );
    expect(token(doc, 'L').x).toBeCloseTo(72, 3);
    const c = token(doc, 'CCC');
    expect((c.x + c.end) / 2).toBeCloseTo(144, 3);
    expect(token(doc, 'RRR').end).toBeCloseTo(216, 3);
    expect(token(doc, '12').end).toBeCloseTo(144, 3);
    expect(token(doc, '12.5').x).toBeCloseTo(token(doc, '12').x, 3);
  });

  it("changes its style's stops: deletes one, adds one, keeps the rest", () => {
    // Word: in a style with stops at 72 and 144, "a", a tab, "b", a tab and "c"
    // put b at 72.0 and c at 144.0; with the 72 deleted and a stop added at
    // 108, b at 108.4 and c at 144.0.
    const styles: ReadonlyArray<DocStyleDef> = [
      { istd: 0, kind: 'paragraph' },
      {
        istd: 1,
        kind: 'paragraph',
        base: 0,
        para: { tabs: { add: [{ pos: 1440 }, { pos: 2880 }] } },
      },
    ];
    const doc = docOf(
      ['a\tb\tc', 'a\tb\tc'],
      [{ istd: 1 }, { istd: 1, tabs: { del: [1440], add: [{ pos: 2160 }] } }],
      styles,
    );
    expect(stopsOf(doc).map((s) => s?.map((t) => t.positionPt))).toEqual([
      [72, 144],
      [108, 144],
    ]);
    const [first, second] = placed(doc);
    expect(first!.filter((t) => t.text !== 'a').map((t) => t.x)).toEqual([72, 144]);
    expect(second!.filter((t) => t.text !== 'a').map((t) => t.x)).toEqual([108, 144]);
  });

  it('carries stops down a chain of styles', () => {
    const styles: ReadonlyArray<DocStyleDef> = [
      { istd: 0, kind: 'paragraph', para: { tabs: { add: [{ pos: 720 }] } } },
      { istd: 1, kind: 'paragraph', base: 0, para: { tabs: { add: [{ pos: 2880, jc: 2 }] } } },
      { istd: 2, kind: 'paragraph', base: 1, para: { tabs: { del: [720] } } },
    ];
    const doc = docOf(['x', 'y', 'z'], [{}, { istd: 1 }, { istd: 2 }], styles);
    expect(stopsOf(doc)).toEqual([
      [{ positionPt: 36, alignment: 'left' }],
      [
        { positionPt: 36, alignment: 'left' },
        { positionPt: 144, alignment: 'right' },
      ],
      [{ positionPt: 144, alignment: 'right' }],
    ]);
  });

  it('deletes the stops within 25 twips of a deletion, or within its own reach', () => {
    // PChgTabsDel: any stop within 25 twips; PChgTabsDelClose: within its
    // rgdxaClose, which is never less than 25.
    const styles: ReadonlyArray<DocStyleDef> = [
      { istd: 0, kind: 'paragraph', para: { tabs: { add: [{ pos: 1440 }, { pos: 2880 }] } } },
    ];
    const doc = docOf(
      ['a', 'b', 'c', 'd'],
      [
        { tabs: { del: [1465] } },
        { tabs: { del: [1466] } },
        { tabs: { del: [1540], close: [100] } },
        { tabs: { del: [1460], close: [5] } },
      ],
      styles,
    );
    expect(stopsOf(doc).map((s) => s?.map((t) => t.positionPt))).toEqual([
      [144],
      [72, 144],
      [144],
      [144],
    ]);
  });

  it('reads on past a tab sprm that deletes, and past one whose size is counted', () => {
    // sprmPChgTabsPapx has a one-byte cb; read as sprmTDefTable's two-byte
    // count, a deletion made it run past the grpprl and the centring after it
    // was lost. sprmPChgTabs' cb of 255 says to count the size from the
    // operand's own counts.
    const doc = docOf(
      ['a', 'b'],
      [
        { tabs: { del: [720], add: [{ pos: 1440 }] }, jc: 1 },
        { tabs: { del: [720], close: [25], add: [{ pos: 2160 }], cb255: true }, jc: 2 },
      ],
    );
    const props = doc.body.map((el) => (el.kind === 'paragraph' ? el.paragraph.properties : {}));
    expect(props.map((p) => [p.alignment, p.tabs?.map((t) => t.positionPt)])).toEqual([
      ['center', [72]],
      ['right', [108]],
    ]);
  });

  it('sets text at a list stop as at a left one, and at no bar stop', () => {
    // Word: with a list (`num`) stop at 72, "ab", a tab and "N" put the N at
    // 72.0, not on the 36pt grid; a bar stop at 36 places nothing, and the B
    // after it goes on to the left stop at 72.
    const doc = docOf(
      ['ab\tN', 'ab\tB'],
      [
        { tabs: { add: [{ pos: 1440, jc: 6 }] } },
        { tabs: { add: [{ pos: 720, jc: 4 }, { pos: 1440 }] } },
      ],
    );
    expect(stopsOf(doc)).toEqual([
      [{ positionPt: 72, alignment: 'left' }],
      [{ positionPt: 72, alignment: 'left' }],
    ]);
    expect(token(doc, 'N').x).toBeCloseTo(72, 3);
    expect(token(doc, 'B').x).toBeCloseTo(72, 3);
  });

  it('gives a paragraph no stops of its own when nothing speaks of tabs', () => {
    const bytes = buildDoc([{ text: 'x\r', compressed: false }], {
      paraRuns: [{ length: 2, jc: 1 }],
    });
    expect(extractDocContent(bytes).paragraphs.map((p) => p.props.tabs)).toEqual([undefined]);
  });
});
