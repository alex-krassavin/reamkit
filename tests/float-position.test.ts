// ECMA-376 Part 1 §17.3.1.11 — where a frame (`w:framePr`) stands when it
// names no vertical anchor, measured against Word (Word for Mac 16, PDF export
// of synthetic probes): it hangs off the MARGIN, as one that says
// `w:vAnchor="margin"` does, wherever its own paragraph happens to be.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** Paragraphs of exactly 12pt lines, so every renderer agrees where they are. */
const P = (inner: string): string =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${inner}</w:p>`;
const T = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
/**
 * The page: A4, its margins 72pt left, 49.25pt right, 57.6pt top, 43.15pt
 * bottom; the header 14.2pt from the top edge and the footer 28.35pt up from
 * the bottom one.
 */
const pgMar =
  '<w:pgSz w:w="11900" w:h="16840"/>' +
  '<w:pgMar w:top="1152" w:right="985" w:bottom="863" w:left="1440" w:header="284" w:footer="567" w:gutter="0"/>';

describe('a frame (§17.3.1.11 w:framePr) keyed into the page', () => {
  /** A 100 × 50pt frame holding `label`, after ten lines of text. */
  const frameDoc = (frames: ReadonlyArray<readonly [label: string, attrs: string]>): Uint8Array =>
    buildDocxFromBody(
      Array.from({ length: 10 }, (_, i) => P(T(`filler ${String(i)}`))).join('') +
        frames
          .map(
            ([label, attrs]) =>
              `<w:p><w:pPr><w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:wrap="around" w:hAnchor="page" ${attrs}/>` +
              `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${T(label)}</w:p>`,
          )
          .join('') +
        `<w:sectPr>${pgMar}</w:sectPr>`,
    );
  const baselines = (docx: Uint8Array): Map<string, number> => {
    const flow = Ream.parse(docx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const out = new Map<string, number>();
    for (const c of laid.pages[0]!.commands) {
      if (c.type !== 'line') continue;
      const text = c.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('');
      out.set(text.trim(), c.baselineY);
    }
    return out;
  };

  it('hangs off the MARGIN when it names no vertical anchor', () => {
    // Word: `w:yAlign="bottom"` with no `w:vAnchor` ends the frame on the
    // margin's foot, as `w:vAnchor="margin"` does; `w:y="1440"` stands 72pt
    // below the margin's top, not 72pt below its paragraph.
    const lines = baselines(
      frameDoc([
        ['NOANCHOR', 'w:x="1440" w:yAlign="bottom"'],
        ['MARGINANCHOR', 'w:x="5040" w:vAnchor="margin" w:yAlign="bottom"'],
        ['TOP', 'w:x="8640" w:yAlign="top"'],
      ]),
    );
    const lines2 = baselines(frameDoc([['OFFSET', 'w:x="1440" w:y="1440"']]));
    expect(lines.get('NOANCHOR')).toBeCloseTo(lines.get('MARGINANCHOR')!, 1);
    // A 50pt frame on the margin's foot (798.85pt down) begins 748.85pt down.
    expect(lines.get('NOANCHOR')! - lines.get('TOP')!).toBeCloseTo(748.85 - 57.6, 1);
    expect(lines2.get('OFFSET')! - lines.get('TOP')!).toBeCloseTo(72, 1);
  });
});
