import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// A floating rect shape anchored to the page: 1828800×914400 EMU = 144×72pt.
const anchoredShape = (posAndWrap: string, attrs = '') =>
  `<w:p><w:r><w:drawing>
    <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ${attrs}>
      <wp:extent cx="1828800" cy="914400"/>
      ${posAndWrap}
      <wp:docPr id="3" name="Float 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>
            </wps:spPr>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing></w:r></w:p>`;

const PAGE_POS =
  '<wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>' +
  '<wp:wrapNone/>';

const TEXT = '<w:p><w:r><w:t>flowing text</w:t></w:r></w:p>';

function layoutOf(docx: Uint8Array) {
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
}

describe('floating drawings (wp:anchor, §20.4.2.3)', () => {
  it('parses the anchor placement onto the shape', () => {
    const { doc } = readDocx(buildDocxFromBody(anchoredShape(PAGE_POS) + TEXT));
    const el = doc.body[0]!;
    if (el.kind !== 'shape') throw new Error('expected shape');
    expect(el.shape.float?.wrap).toBe('none');
    expect(el.shape.float?.posH).toEqual({ relativeFrom: 'page', offsetPt: 100 });
    expect(el.shape.float?.posV).toEqual({ relativeFrom: 'page', offsetPt: 50 });
  });

  it('renders out of flow at the page-relative position (text unaffected)', () => {
    const laid = layoutOf(buildDocxFromBody(anchoredShape(PAGE_POS) + TEXT));
    const withoutFloat = layoutOf(buildDocxFromBody(TEXT));
    const cmds = laid.pages[0]!.commands;
    const shape = cmds.find((c) => c.type === 'shape') as unknown as {
      shape: { transform: ReadonlyArray<number> };
    };
    expect(shape).toBeDefined();
    // 1270000 EMU = 100pt from the page's left edge.
    expect(shape.shape.transform[4]).toBeCloseTo(100, 1);
    // The text line sits exactly where it would without the float.
    const lineY = (cs: ReadonlyArray<{ type: string }>) =>
      (cs.find((c) => c.type === 'line') as unknown as { baselineY: number }).baselineY;
    expect(lineY(cmds)).toBeCloseTo(lineY(withoutFloat.pages[0]!.commands), 4);
  });

  it('behindDoc sinks below body text; default floats above it', () => {
    const behind = layoutOf(buildDocxFromBody(anchoredShape(PAGE_POS, 'behindDoc="1"') + TEXT))
      .pages[0]!.commands;
    const bShape = behind.findIndex((c) => c.type === 'shape');
    const bLine = behind.findIndex((c) => c.type === 'line');
    expect(bShape).toBeLessThan(bLine);

    const front = layoutOf(buildDocxFromBody(anchoredShape(PAGE_POS) + TEXT)).pages[0]!.commands;
    expect(front.findIndex((c) => c.type === 'shape')).toBeGreaterThan(
      front.findIndex((c) => c.type === 'line'),
    );
  });

  it('margin-relative vertical offsets hang off the top margin', () => {
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>' +
      '<wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapNone/>';
    const laid = layoutOf(buildDocxFromBody(anchoredShape(pos) + TEXT));
    const shape = laid.pages[0]!.commands.find((c) => c.type === 'shape') as unknown as {
      shape: { transform: ReadonlyArray<number> };
    };
    // A4 + 1" margins: content right edge at 523.3pt; 144pt wide → x = 379.3.
    expect(shape.shape.transform[4]).toBeCloseTo(595.276 - 72 - 144, 0);
  });

  it('side-wrapping floats stay in flow (square ≈ block, v1)', () => {
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapSquare wrapText="bothSides"/>';
    const laid = layoutOf(buildDocxFromBody(anchoredShape(pos) + TEXT));
    const withoutFloat = layoutOf(buildDocxFromBody(TEXT));
    // In-flow: the text moves DOWN relative to the no-float document.
    const lineY = (cs: ReadonlyArray<{ type: string }>) =>
      (cs.find((c) => c.type === 'line') as unknown as { baselineY: number }).baselineY;
    expect(lineY(laid.pages[0]!.commands)).toBeGreaterThan(lineY(withoutFloat.pages[0]!.commands));
  });
});
