// E-DOCX WT3 (OfficeMath) — math survives docx → FlowDoc → docx. The omml
// serializer is the inverse of the parser, so a fraction + superscript come
// back with the same node structure and literal symbols.

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement, MathNode } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';

const M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

// a/b · √(c) · x²  — exercises fraction, radical and superscript.
const MATH =
  `<m:oMath ${M_NS}>` +
  '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' +
  '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>c</m:t></m:r></m:e></m:rad>' +
  '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>' +
  '</m:oMath>';

const findMath = (body: ReadonlyArray<BodyElement>): MathNode | undefined => {
  for (const el of body) {
    if (el.kind !== 'paragraph') continue;
    for (const run of el.paragraph.runs) if (run.math) return run.math;
  }
  return undefined;
};

describe('docx OfficeMath write-back (WT3)', () => {
  it('round-trips a fraction / radical / superscript through docx → FlowDoc → docx', async () => {
    const docx = buildDocxFromBody(`<w:p>${MATH}</w:p>`);
    expect(findMath(Ream.parse(docx).flow.body)).toBeDefined();

    const { bytes, losses } = await Ream.parse(docx).convertWithReport('docx');
    expect(losses.some((l) => /math/i.test(l.detail))).toBe(false);

    const math = findMath(Ream.parse(bytes).flow.body);
    expect(math).toBeDefined();
    const dump = JSON.stringify(math);
    expect(dump).toContain('"fraction"');
    expect(dump).toContain('"radical"');
    expect(dump).toContain('"script"');
    for (const sym of ['"a"', '"b"', '"c"', '"x"', '"2"']) expect(dump).toContain(sym);
  });
});
