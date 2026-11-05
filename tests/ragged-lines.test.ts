// A line is drawn no wider than its measure. The breaker weighs a line knowing
// how far its spaces may be squeezed (GLUE_SHRINK_RATIO), and only a justified
// line is drawn squeezed — so a ragged paragraph offered the squeeze kept words
// its lines were too short for and ran them past the margin: the Turkic probes
// (.git/wip-turkic/notes.md) had 4 to 25 such lines a document, a Chuvash
// sentence in Courier New 13pt into the margin, where Word has none.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { Line } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { computeJustifyExtra } from '@/layout/line-paint';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = { regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')) };

/** Words of 1 to 12 letters, the same every run (a linear congruential stream). */
function words(seed: number, n: number): string {
  let x = seed;
  const next = (): number => (x = (x * 1103515245 + 12345) % 2147483648) / 2147483648;
  return Array.from({ length: n }, () =>
    Array.from({ length: 1 + Math.floor(next() * 12) }, () =>
      String.fromCharCode(97 + Math.floor(next() * 26)),
    ).join(''),
  ).join(' ');
}

/** Each line laid out, with how far it reaches when drawn: its glyphs and its spaces as squeezed or stretched. */
function drawnLines(jc: string): Array<{ reach: number; measure: number; last: boolean }> {
  const body = Array.from(
    { length: 40 },
    (_, k) =>
      `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr><w:r><w:t xml:space="preserve">${words(k + 1, 60)}</w:t></w:r></w:p>`,
  ).join('');
  const flow = Ream.parse(buildDocxFromBody(body)).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out: Array<{ reach: number; measure: number; last: boolean }> = [];
  for (const page of laid.pages) {
    for (const c of page.commands) {
      if (c.type !== 'line') continue;
      const line = (c as unknown as { line: Line }).line;
      let last = line.tokens.length - 1;
      while (last >= 0 && line.tokens[last]!.kind === 'text' && line.tokens[last]!.isSpace) last--;
      const extra = computeJustifyExtra(line);
      let reach = 0;
      for (let i = 0; i <= last; i++) {
        const t = line.tokens[i]!;
        reach += t.widthPt + (t.kind === 'text' && t.isSpace ? extra : 0);
      }
      out.push({ reach, measure: line.availableWidthPt, last: line.isLastInParagraph });
    }
  }
  return out;
}

describe('a line drawn no wider than its measure', () => {
  for (const jc of ['left', 'center', 'right', 'both', 'distribute']) {
    it(`holds for ${jc}-aligned text`, () => {
      const lines = drawnLines(jc);
      expect(lines.length).toBeGreaterThan(100);
      const over = lines.filter((l) => l.reach > l.measure + 0.01);
      expect(over).toEqual([]);
    });
  }

  it('leaves a ragged line its spaces as they are', () => {
    // Nothing is squeezed or stretched where the text is not justified.
    for (const l of drawnLines('left')) expect(l.reach).toBeLessThanOrEqual(l.measure + 0.01);
  });
});
