import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc } from './fixtures/build-doc';
import { buildDocxFromBody } from './fixtures/build-docx';
import type { FlowDoc } from '@/core/ir/flow';
import type { Line, TextLineItem } from '@/layout/page-doc';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDoc } from '@/word/doc/doc-reader';
import { readDocx } from '@/word/docx-reader';
import { writeDocx } from '@/word/docx-writer';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** Every line of the document, as laid out. */
function lines(flow: FlowDoc): Array<Line> {
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.flatMap((p) =>
    p.commands.filter((c): c is TextLineItem => c.type === 'line').map((c) => c.line),
  );
}

/** Where each tab on a line ENDS, measured from the text margin. */
function tabEnds(line: Line): Array<number> {
  let x = line.resolved.indentLeft + (line.firstLine ? line.resolved.indentFirstLine : 0);
  const ends: Array<number> = [];
  for (const t of line.tokens) {
    x += t.widthPt;
    if (t.kind === 'text' && t.tab === true) ends.push(Math.round(x * 100) / 100);
  }
  return ends;
}

/** A paragraph of `text`, a `\t` in it a `w:tab`. */
function para(text: string, pPr = ''): string {
  const runs = text
    .split('\t')
    .map(
      (part, i) =>
        (i > 0 ? '<w:r><w:tab/></w:r>' : '') + (part ? `<w:r><w:t>${part}</w:t></w:r>` : ''),
    )
    .join('');
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
}

const BODY = [para('P1\tA'), para('\t'.repeat(10) + 'B'), para('abcdefg\tC')].join('');

// Measured in Word for Mac (compatibility mode 15) on a probe of these very
// paragraphs: with `w:defaultTabStop w:val="708"` the first tab ends at 35.4pt,
// ten of them at 354pt, and one after seven letters at 70.8pt — and twelve
// tabs and 42pt of Courier still fit a 468pt line that a 720-twip grid wraps.
// LibreOffice sets them on the same grid.
describe('§17.15.1.25 w:defaultTabStop — the automatic tab stops', () => {
  it('lays the stops at the interval the document states', () => {
    const flow = readDocx(
      buildDocxFromBody(BODY, { settingsXml: '<w:defaultTabStop w:val="708"/>' }),
    ).doc;
    expect(flow.defaultTabStopPt).toBeCloseTo(35.4, 6);
    const [p1, ten, after] = lines(flow);
    expect(tabEnds(p1!)).toEqual([35.4]);
    expect(tabEnds(ten!)).toEqual(
      Array.from({ length: 10 }, (_, i) => Math.round((i + 1) * 3540) / 100),
    );
    expect(tabEnds(after!)).toEqual([70.8]);
  });

  it('keeps the half inch where the document states none', () => {
    const flow = readDocx(buildDocxFromBody(BODY)).doc;
    expect(flow.defaultTabStopPt).toBeUndefined();
    const [p1, ten, after] = lines(flow);
    expect(tabEnds(p1!)).toEqual([36]);
    expect(tabEnds(ten!).at(-1)).toBe(360);
    expect(tabEnds(after!)).toEqual([72]);
  });

  it('counts the stops from the margin, not the indent, and only past the own stops', () => {
    // Word: "D" at 35.4 under a 25pt indent, not at 60.4; past an own stop at
    // 50pt the next tab takes the grid's 70.8.
    const body =
      para('\tD', '<w:ind w:left="500"/>') +
      para('\t\tE', '<w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs>');
    const flow = readDocx(
      buildDocxFromBody(body, { settingsXml: '<w:defaultTabStop w:val="708"/>' }),
    ).doc;
    const [d, e] = lines(flow);
    expect(tabEnds(d!)).toEqual([35.4]);
    expect(tabEnds(e!)).toEqual([50, 70.8]);
  });

  it("sends a list marker's tab past a narrow hanging indent to the document's grid", () => {
    // The ru-locale numbering that started this: a marker wider than its
    // 284-twip hanging indent. Word sets the text at 35.4pt after it.
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>' +
      '<w:lvlText w:val="ABCD"/><w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>' +
      '<w:rPr><w:sz w:val="20"/></w:rPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const body = para('F', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    const flow = readDocx(
      buildDocxFromBody(body, { numberingXml, settingsXml: '<w:defaultTabStop w:val="708"/>' }),
    ).doc;
    expect(tabEnds(lines(flow)[0]!)).toEqual([35.4]);
  });

  it('gives a tab no width where the interval is zero', () => {
    // Word and LibreOffice alike: "P1", a tab, "A" reads "P1A".
    const flow = readDocx(
      buildDocxFromBody(para('P1\tA'), { settingsXml: '<w:defaultTabStop w:val="0"/>' }),
    ).doc;
    const tab = lines(flow)[0]!.tokens.find((t) => t.kind === 'text' && t.tab === true);
    expect(tab?.widthPt).toBe(0);
  });

  it('reads a measure written with a unit', () => {
    const flow = readDocx(
      buildDocxFromBody(para('P1\tA'), { settingsXml: '<w:defaultTabStop w:val="1cm"/>' }),
    ).doc;
    expect(flow.defaultTabStopPt).toBeCloseTo(72 / 2.54, 6);
  });

  it('writes the interval back into settings.xml', () => {
    const flow = readDocx(
      buildDocxFromBody(BODY, { settingsXml: '<w:defaultTabStop w:val="708"/>' }),
    ).doc;
    const bytes = writeDocx(flow).bytes;
    const settings = OpcPackage.open(bytes).getPart('word/settings.xml');
    expect(new TextDecoder().decode(settings)).toContain('<w:defaultTabStop w:val="708"/>');
    expect(readDocx(bytes).doc.defaultTabStopPt).toBeCloseTo(35.4, 6);
  });

  it("takes a .doc's interval from Dop.dxaTab", () => {
    // LibreOffice writes a .doc's dxaTab of 708 (52420.doc) as
    // `w:defaultTabStop w:val="708"` when it saves the file as a .docx.
    const doc = (dxaTab?: number): FlowDoc =>
      readDoc(
        buildDoc([{ text: 'P1\tA\r', compressed: false }], {
          ...(dxaTab !== undefined ? { dxaTab } : {}),
        }),
      ).doc;
    const flow = doc(708);
    expect(flow.defaultTabStopPt).toBeCloseTo(35.4, 6);
    expect(tabEnds(lines(flow)[0]!)).toEqual([35.4]);
    const plain = doc();
    expect(plain.defaultTabStopPt).toBeUndefined();
    expect(tabEnds(lines(plain)[0]!)).toEqual([36]);
  });
});
