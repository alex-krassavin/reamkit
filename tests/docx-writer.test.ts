import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { OpcPackage } from '@/core/opc';
import { writeDocx } from '@/word/docx-writer';
import { readDocx } from '@/word/docx-reader';

const decode = (b: Uint8Array) => new TextDecoder().decode(b);

const BODY =
  '<w:p><w:r><w:t>first paragraph</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">spaced  &amp; escaped &lt;text&gt;</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/></w:sectPr>';

describe('docx writer (E-DOCX D2 skeleton)', () => {
  it('round-trips text and page geometry through its own reader', () => {
    const { doc: flow } = readDocx(buildDocxFromBody(BODY));
    const { bytes, losses } = writeDocx(flow);
    expect(losses).toHaveLength(0);

    const { doc: again } = readDocx(bytes);
    const texts = again.body.flatMap((el) =>
      el.kind === 'paragraph' ? [el.paragraph.runs.map((r) => r.text).join('')] : [],
    );
    expect(texts).toEqual(['first paragraph', 'spaced  & escaped <text>']);
    // Page geometry survives (A4 portrait, custom margins).
    const sect = again.sections[0]?.properties ?? again.section;
    expect(sect?.pageSize?.width).toBeCloseTo(11906 / 20, 1);
    expect(sect?.margins?.left).toBeCloseTo(54, 1);
  });

  it('is exposed as the docx target on Ream.convert', async () => {
    const src = buildDocxFromBody('<w:p><w:r><w:t>via facade</w:t></w:r></w:p>');
    const out = await Ream.parse(src).convert('docx');
    const pkg = OpcPackage.open(out);
    expect(decode(pkg.getMainDocument().data)).toContain('via facade');
  });

  it('is deterministic', () => {
    const { doc: flow } = readDocx(buildDocxFromBody(BODY));
    expect([...writeDocx(flow).bytes]).toEqual([...writeDocx(flow).bytes]);
  });

  it('reports unwritten block kinds as losses (v0)', () => {
    const tbl =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc: flow } = readDocx(buildDocxFromBody(tbl));
    const { losses } = writeDocx(flow);
    expect(losses.some((l) => l.feature === 'tables' && l.severity === 'dropped')).toBe(true);
  });

  it('strict mode surfaces writer losses', async () => {
    const tbl =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    await expect(
      Ream.parse(buildDocxFromBody(tbl)).convert('docx', { strict: true }),
    ).rejects.toThrow();
  });
});
