// §18.18.3 ST_BorderStyle — the names are PATTERNS as well as weights, and a
// cell that asks for a dashed medium rule or a double one has to get it.
// 59264.xlsx is a sampler of every style in the enumeration; it drew its
// MEDIUM_DASHED family solid and its DOUBLE as a single hairline.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { Border } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';

// One border per style, each on its own cellXf, so cell N of row 0 carries
// style N. `<border>` order matches STYLES below.
const NAMES = [
  'thin',
  'medium',
  'thick',
  'double',
  'dashed',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
] as const;

const STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="${String(NAMES.length + 1)}">
    <border/>
    ${NAMES.map((n) => `<border><top style="${n}"/></border>`).join('')}
  </borders>
  <cellXfs count="${String(NAMES.length + 1)}">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    ${NAMES.map(
      (_, i) =>
        `<xf numFmtId="0" fontId="0" fillId="0" borderId="${String(i + 1)}" applyBorder="1"/>`,
    ).join('')}
  </cellXfs>`;

function topBorders(): Array<Border | undefined> {
  const flow = Ream.parse(
    buildXlsx({
      rows: [NAMES.map((n, i) => ({ value: n, styleIndex: i + 1 }))],
      stylesXml: STYLES,
    }),
  ).flow;
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a grid table');
  return (table.table.rows[0]?.cells ?? []).map((c) => c.properties.borders?.top);
}

describe('border styles (§18.18.3)', () => {
  it('keeps the pattern of a medium dashed rule, at its medium weight', () => {
    const got = topBorders();
    const by = (name: (typeof NAMES)[number]): Border | undefined => got[NAMES.indexOf(name)];
    // Plain medium is a solid 1.5pt rule…
    expect(by('medium')).toMatchObject({ style: 'single', width: 1.5 });
    // …and every medium PATTERN keeps the weight and gains the dashes.
    for (const n of ['mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot'] as const) {
      expect(by(n)).toMatchObject({ style: 'dashed', width: 1.5 });
    }
    // A thin dashed rule stays thin.
    expect(by('dashed')).toMatchObject({ style: 'dashed', width: 0.75 });
  });

  it('gives a double rule the width its two lines and their gap need', () => {
    const by = (name: (typeof NAMES)[number]): Border | undefined =>
      topBorders()[NAMES.indexOf(name)];
    // Three screen pixels across: a line, a gap and a line. At 0.75pt there is
    // no room for the gap and the pair renders as one hairline.
    expect(by('double')).toMatchObject({ style: 'double', width: 2.25 });
  });
});
