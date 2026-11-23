// ECMA-376 Part 1 §17.3.1.37 / §17.3.3.32 — tabs are settled one after
// another, each from where the one before it left the pen: a right, centre or
// decimal stop with nothing to place before the next tab still takes the pen
// to it, and the next tab goes on to the stop after it (issue #146). Every
// expectation was measured in Word for Mac on Courier New probes
// (compatibility mode 15, character positions by AppleScript); LibreOffice
// sets them the same, except where noted.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { Line, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** The issue's page: A4, 1440 twips left, 985 right — a 473.75pt measure. */
const PAGE =
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="985" w:bottom="1440" w:left="1440" w:header="709" w:footer="709"/>';

type Stop = readonly [kind: string, twips: number];

const tabsOf = (stops: ReadonlyArray<Stop>): string =>
  `<w:tabs>${stops.map(([v, pos]) => `<w:tab w:val="${v}" w:pos="${pos}"/>`).join('')}</w:tabs>`;

/** A paragraph of `text`, each `\t` in it a `w:tab`, against `stops`. */
const para = (text: string, stops: ReadonlyArray<Stop>, pPr = ''): string => {
  const runs = text
    .split('\t')
    .map(
      (part, i) =>
        (i > 0 ? '<w:r><w:tab/></w:r>' : '') +
        (part ? `<w:r><w:t xml:space="preserve">${part}</w:t></w:r>` : ''),
    )
    .join('');
  return `<w:p><w:pPr>${tabsOf(stops)}${pPr}</w:pPr>${runs}</w:p>`;
};

function lines(body: string, footerXml?: string): Array<Line> {
  const section =
    footerXml === undefined
      ? `<w:sectPr>${PAGE}</w:sectPr>`
      : `<w:sectPr><w:footerReference w:type="default" r:id="rId11"/>${PAGE}</w:sectPr>`;
  const flow = Ream.parse(
    buildDocxFromBody(body + section, footerXml === undefined ? {} : { footerXml }),
  ).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.flatMap((p) =>
    p.commands.filter((c): c is TextLineItem => c.type === 'line').map((c) => c.line),
  );
}

/** Where each printing piece of a line starts and ends, from the text margin. */
function pieces(line: Line): Array<{ text: string; from: number; to: number }> {
  let x = line.resolved.indentLeft + (line.firstLine ? line.resolved.indentFirstLine : 0);
  const out: Array<{ text: string; from: number; to: number }> = [];
  for (const t of line.tokens) {
    if (t.kind === 'text' && !t.isSpace) out.push({ text: t.text, from: x, to: x + t.widthPt });
    x += t.widthPt;
  }
  return out;
}

/** The piece `text` of the one line of a one-paragraph document. */
function piece(
  text: string,
  stops: ReadonlyArray<Stop>,
  needle: string,
): { from: number; to: number } {
  const [line] = lines(para(text, stops));
  const hit = pieces(line!).find((p) => p.text === needle);
  if (!hit) throw new Error(`no ${needle} on the line`);
  return hit;
}

const C = 'center';
const R = 'right';
const D = 'decimal';

describe('a right, centre or decimal stop with nothing before the next tab (issue #146)', () => {
  it("ends the issue's page number at the right stop, past the centre one", () => {
    // Word: the first tab ends at the centre stop (8222 twips, 411.1pt), the
    // second carries "8" to end at the right one (9356, 467.8).
    const stops: ReadonlyArray<Stop> = [
      ['clear', 4513],
      ['clear', 9026],
      [C, 8222],
      [R, 9356],
    ];
    const at = piece('Site name [street address] - [Nov 2026]\t\t8', stops, '8');
    expect(at.to).toBeCloseTo(467.8, 3);
  });

  it('does so in the footer, around a PAGE field', () => {
    const footer =
      `<w:p><w:pPr><w:pStyle w:val="Footer"/>${tabsOf([
        ['clear', 4513],
        ['clear', 9026],
        [C, 8222],
        [R, 9356],
      ])}</w:pPr>` +
      '<w:r><w:t>Site name [street address] - [Nov 2026]</w:t></w:r>' +
      '<w:r><w:tab/></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>8</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const footerLine = lines(para('body', []), footer).find((l) =>
      l.tokens.some((t) => t.kind === 'text' && t.text.startsWith('Site')),
    );
    const number = pieces(footerLine!).at(-1)!;
    expect(number.text).toBe('1');
    expect(number.to).toBeCloseTo(467.8, 3);
  });

  it('takes the pen to the stop at the start of the line and mid-line alike', () => {
    // Word: "XYZ" ends at 400pt after a centre stop at 200, a right one, a
    // decimal one; from the start of the line and after "ab".
    for (const first of [C, R, D]) {
      expect(
        piece(
          '\t\tXYZ',
          [
            [first, 4000],
            [R, 8000],
          ],
          'XYZ',
        ).to,
      ).toBeCloseTo(400, 3);
      expect(
        piece(
          'ab\t\tXYZ',
          [
            [first, 4000],
            [R, 8000],
          ],
          'XYZ',
        ).to,
      ).toBeCloseTo(400, 3);
    }
  });

  it('sets the next segment by the stop after: left, centre or decimal', () => {
    // Word: "XYZ" at 299.85 for a left stop at 300; centred on a centre stop
    // at 400; "1.5" with its point at 400 on a decimal one.
    expect(
      piece(
        'ab\t\tXYZ',
        [
          [C, 4000],
          ['left', 6000],
        ],
        'XYZ',
      ).from,
    ).toBeCloseTo(300, 3);
    const centred = piece(
      'ab\t\tXYZ',
      [
        [C, 4000],
        [C, 8000],
      ],
      'XYZ',
    );
    expect((centred.from + centred.to) / 2).toBeCloseTo(400, 3);
    const [line] = lines(
      para('ab\t\t1.5', [
        [C, 4000],
        [D, 8000],
      ]),
    );
    const num = pieces(line!).find((p) => p.text === '1.5')!;
    const one = line!.tokens.find((t) => t.kind === 'text' && t.text === '1.5');
    const intPart =
      one?.kind === 'text' ? one.font.measure.textWidthPt('1', one.fontSizePt) : Number.NaN;
    expect(num.from + intPart).toBeCloseTo(400, 3);
  });

  it('passes each stop once, three tabs in a row', () => {
    // Word: centre 200, right 300, right 400 — "XYZ" ends at 400.
    const at = piece(
      'ab\t\t\tXYZ',
      [
        [C, 4000],
        [R, 6000],
        [R, 8000],
      ],
      'XYZ',
    );
    expect(at.to).toBeCloseTo(400, 3);
  });

  it('goes on to the default grid past the last stop', () => {
    // Word: after a lone centre or right stop at 200 the next tab reaches the
    // half-inch grid at 216, the one after that 252; text past the centre stop
    // sends the first tab to the right stop at 400 and the second to 432.
    expect(piece('ab\t\tXYZ', [[C, 4000]], 'XYZ').from).toBeCloseTo(216, 3);
    expect(piece('ab\t\t\tXYZ', [[C, 4000]], 'XYZ').from).toBeCloseTo(252, 3);
    expect(piece('ab\t\tXYZ', [[R, 4000]], 'XYZ').from).toBeCloseTo(216, 3);
    expect(
      piece(
        `${'A'.repeat(35)}\t\tXYZ`,
        [
          [C, 4000],
          [R, 8000],
        ],
        'XYZ',
      ).from,
    ).toBeCloseTo(432, 3);
    expect(
      piece(
        '\t\t\tXYZ',
        [
          [C, 4000],
          [R, 8000],
        ],
        'XYZ',
      ).from,
    ).toBeCloseTo(432, 3);
  });

  it('hands a stop on when the next tab has nowhere left to go on the line', () => {
    // Word moves that tab down and ends "XYZ" at the margin on the next line;
    // the line here is already set, so "XYZ" keeps the margin, one line up.
    const at = piece('ab\t\tXYZ', [[R, 9475]], 'XYZ');
    expect(at.to).toBeCloseTo(473.75, 3);
  });
});

describe("a right or centre segment's trailing spaces (§17.3.1.37)", () => {
  it('hang past the stop', () => {
    // Word: "XY" and two spaces end the "XY" at a right stop at 200, and
    // centre the "XY" on a centre one; LibreOffice counts the spaces.
    const right = piece(
      'ab\tXY  \tQQ',
      [
        [R, 4000],
        [R, 8000],
      ],
      'XY',
    );
    expect(right.to).toBeCloseTo(200, 3);
    const centre = piece(
      'ab\tXY  \tQQ',
      [
        [C, 4000],
        [R, 8000],
      ],
      'XY',
    );
    expect((centre.from + centre.to) / 2).toBeCloseTo(200, 3);
  });

  it('count when they lead', () => {
    // Word: two spaces and "XY" end the "XY" at 200, the spaces before it.
    const lead = piece(
      'ab\t  XY\tQQ',
      [
        [R, 4000],
        [R, 8000],
      ],
      'XY',
    );
    expect(lead.to).toBeCloseTo(200, 3);
  });
});
