import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
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
    // Letter (the page a docx with no `w:pgSz` is on) + 1" margins: content
    // right edge at 540pt; 144pt wide → x = 396.
    expect(shape.shape.transform[4]).toBeCloseTo(612 - 72 - 144, 0);
  });

  it('side-wrapping floats keep the text at its height (no vertical push)', () => {
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapSquare wrapText="bothSides"/>';
    const laid = layoutOf(buildDocxFromBody(anchoredShape(pos) + TEXT));
    const withoutFloat = layoutOf(buildDocxFromBody(TEXT));
    // Out of flow now: the first text line keeps its baseline and flows
    // BESIDE the float instead of below it.
    const first = (cs: ReadonlyArray<{ type: string }>) =>
      cs.find((c) => c.type === 'line') as unknown as { baselineY: number; originX: number };
    expect(first(laid.pages[0]!.commands).baselineY).toBeCloseTo(
      first(withoutFloat.pages[0]!.commands).baselineY,
      1,
    );
    expect(first(laid.pages[0]!.commands).originX).toBeGreaterThan(200);
  });

  it('wrapSquare narrows lines beside the float and restores full width below', () => {
    // 144x72pt float at the left margin, paragraph-anchored; long text flows.
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapSquare wrapText="bothSides"/>';
    const words = Array.from({ length: 120 }, (_, i) => `word${i}`).join(' ');
    const docx = buildDocxFromBody(
      anchoredShape(pos) + `<w:p><w:r><w:t>${words}</w:t></w:r></w:p>`,
    );
    const laid = layoutOf(docx);
    const cmds = laid.pages[0]!.commands;
    const lines = cmds.filter((c) => c.type === 'line') as unknown as Array<{
      originX: number;
      baselineY: number;
    }>;
    expect(lines.length).toBeGreaterThan(4);
    // A4 + 1" margins: marginLeft = 72. The float spans x 72..216, y-down top
    // of the paragraph. Lines BESIDE the float must start right of it.
    const floatBottomYDown = 72 + 72; // page top margin + float height? — derive instead:
    const beside = lines.filter((l) => l.originX > 200);
    const below = lines.filter((l) => l.originX < 100);
    expect(beside.length).toBeGreaterThan(0); // narrowed, shifted lines exist
    expect(below.length).toBeGreaterThan(0); // and full-width lines resume
    // Every shifted line sits ABOVE every resumed line (y-down: smaller y).
    const maxBesideY = Math.max(...beside.map((l) => l.baselineY));
    const minBelowY = Math.min(...below.map((l) => l.baselineY));
    expect(maxBesideY).toBeLessThan(minBelowY);
    // The float itself renders (front layer) as a shape command.
    expect(cmds.some((c) => c.type === 'shape')).toBe(true);
  });

  it('fills BOTH gaps beside a float at one baseline (§20.4.2.3 bothSides)', () => {
    // A 144pt box in the MIDDLE of the measure: text stands to its left and
    // carries on to its right at the same height. Taking the wider side alone
    // left fdo65718.docx's narrow column blank down the whole float.
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:posOffset>1524000</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapSquare wrapText="bothSides"/>';
    const words = Array.from({ length: 120 }, (_, i) => `w${String(i)}`).join(' ');
    const laid = layoutOf(
      buildDocxFromBody(anchoredShape(pos) + `<w:p><w:r><w:t>${words}</w:t></w:r></w:p>`),
    );
    const lines = laid.pages[0]!.commands.filter((c) => c.type === 'line').map(
      (c) => c as unknown as { originX: number; baselineY: number },
    );
    // Two segments share a baseline, one each side of the box.
    const byBaseline = new Map<number, Array<number>>();
    for (const l of lines) {
      const at = byBaseline.get(Math.round(l.baselineY)) ?? [];
      at.push(l.originX);
      byBaseline.set(Math.round(l.baselineY), at);
    }
    const pairs = [...byBaseline.values()].filter((xs) => xs.length === 2);
    expect(pairs.length).toBeGreaterThan(0);
    for (const [a, b] of pairs) expect(Math.abs(a! - b!)).toBeGreaterThan(144);
  });

  it('keeps one side when the anchor names one (§20.4.2.3 wrapText="right")', () => {
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:posOffset>1524000</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapSquare wrapText="right"/>';
    const words = Array.from({ length: 120 }, (_, i) => `w${String(i)}`).join(' ');
    const laid = layoutOf(
      buildDocxFromBody(anchoredShape(pos) + `<w:p><w:r><w:t>${words}</w:t></w:r></w:p>`),
    );
    const baselines = laid.pages[0]!.commands.filter((c) => c.type === 'line').map((c) =>
      Math.round((c as unknown as { baselineY: number }).baselineY),
    );
    expect(new Set(baselines).size).toBe(baselines.length);
  });

  it('side-wrapped floats no longer consume vertical flow space', () => {
    const pos =
      '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:wrapSquare wrapText="bothSides"/>';
    const text = '<w:p><w:r><w:t>short text</w:t></w:r></w:p>';
    const withFloat = layoutOf(buildDocxFromBody(anchoredShape(pos) + text));
    const lines = withFloat.pages[0]!.commands.filter(
      (c) => c.type === 'line',
    ) as unknown as Array<{
      originX: number;
    }>;
    // The single short line sits beside the float (shifted), not under it.
    expect(lines[0]!.originX).toBeGreaterThan(200);
  });
});

// A one-cell table whose only paragraph carries an anchored shape (§20.4.2.4).
const cellFloat = (posAndWrap: string, attrs = '') =>
  `<w:tbl>
    <w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
    <w:tr><w:tc>
      <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
      ${anchoredShape(posAndWrap, attrs)}
    </w:tc></w:tr>
  </w:tbl>${TEXT}`;

const shapeX = (laid: ReturnType<typeof layoutOf>): number => {
  const shape = laid.pages[0]!.commands.find((c) => c.type === 'shape');
  if (!shape) throw new Error('the anchored shape was not drawn');
  return (shape as unknown as { shape: { transform: ReadonlyArray<number> } }).shape.transform[4]!;
};

describe('a drawing anchored inside a table cell (§20.4.2.4)', () => {
  // The table starts at the left margin (72pt) and the cell keeps Word's
  // default 5.4pt of padding, so a cell-relative offset lands 77.4pt further
  // right than the same offset read against the page.
  const AT_100 =
    '<wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>' +
    '<wp:wrapNone/>';

  it('is drawn at all — it used to be dropped for having nowhere to stand', () => {
    const laid = layoutOf(buildDocxFromBody(cellFloat(AT_100)));
    expect(laid.pages[0]!.commands.some((c) => c.type === 'shape')).toBe(true);
  });

  it('measures its position in the CELL, which is what layoutInCell means', () => {
    expect(shapeX(layoutOf(buildDocxFromBody(cellFloat(AT_100))))).toBeCloseTo(177.4, 0);
  });

  it('…and reaches past the table to the page when layoutInCell is off', () => {
    const laid = layoutOf(buildDocxFromBody(cellFloat(AT_100, 'layoutInCell="0"')));
    expect(shapeX(laid)).toBeCloseTo(100, 0);
  });

  it('takes no room in the cell it is anchored in', () => {
    const withFloat = layoutOf(buildDocxFromBody(cellFloat(AT_100)));
    const bare = layoutOf(
      buildDocxFromBody(cellFloat(AT_100).replace(anchoredShape(AT_100), '<w:p/>')),
    );
    const lineY = (l: ReturnType<typeof layoutOf>) =>
      (
        l.pages[0]!.commands.filter((c) => c.type === 'line').at(-1) as unknown as {
          baselineY: number;
        }
      ).baselineY;
    expect(lineY(withFloat)).toBeCloseTo(lineY(bare), 4);
  });
});

// Two boxes of one chain: the first holds the words (`wps:txbx id`), the second
// only says it continues them (`wps:linkedTxbx id seq`).
const chained = (heightEmu: number, text: string) =>
  buildDocxFromBody(
    `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="1828800" cy="${String(heightEmu)}"/>
        <wp:positionH relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>
        <wp:wrapNone/>
        <wp:docPr id="1" name="Box 1"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="${String(heightEmu)}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
              <wps:txbx id="7"><w:txbxContent>${text}</w:txbxContent></wps:txbx>
              <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing></w:r></w:p>` +
      `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="1828800" cy="3657600"/>
        <wp:positionH relativeFrom="page"><wp:posOffset>3175000</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>
        <wp:wrapNone/>
        <wp:docPr id="2" name="Box 2"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="3657600"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
              <wps:linkedTxbx id="7" seq="1"/>
              <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing></w:r></w:p>`,
  );

const PARAS = Array.from(
  { length: 8 },
  (_, i) => `<w:p><w:r><w:t>line ${String(i)}</w:t></w:r></w:p>`,
).join('');

// The x of every line that carries words, by the text on it — the two body
// paragraphs the drawings hang off draw an empty line each.
const linesByText = (laid: ReturnType<typeof layoutOf>) =>
  laid.pages[0]!.commands.filter((c) => c.type === 'line')
    .map(
      (c) =>
        c as unknown as { originX: number; line: { tokens: ReadonlyArray<{ text?: string }> } },
    )
    .map((c) => ({ x: c.originX, text: c.line.tokens.map((t) => t.text ?? '').join('') }))
    .filter((d) => d.text !== '');

describe('linked text boxes (wps:linkedTxbx)', () => {
  it('continues what will not fit in the next box of the chain', () => {
    // A box 36pt tall holds three 12pt lines; the rest belong to box 2, which
    // stands 200pt to its right.
    const laid = layoutOf(chained(457200, PARAS));
    const drawn = linesByText(laid);
    expect(drawn.map((d) => d.text)).toEqual(PARAS.match(/line \d/gu));
    const first = drawn[0]!.x;
    const carried = drawn.filter((d) => d.x > first + 100);
    expect(carried.length).toBeGreaterThan(0);
    expect(drawn.filter((d) => d.x === first).length).toBeGreaterThan(0);
    // Every carried line comes AFTER every kept one: one chain, read in order.
    const lastKept = drawn.map((d) => d.x === first).lastIndexOf(true);
    expect(lastKept).toBeLessThan(drawn.indexOf(carried[0]!));
  });

  it('leaves a box that holds its whole text alone', () => {
    const laid = layoutOf(chained(3657600, PARAS));
    const drawn = linesByText(laid);
    const first = drawn[0]!.x;
    expect(drawn.every((d) => d.x === first)).toBe(true);
  });
});

describe('a tiled picture fill (§14.1.2.5 type="tile" / §20.1.8.58 a:tile)', () => {
  // A 2x2 PNG is 1.5pt square at the 96 dpi Office measures a picture in, so a
  // 144x72pt shape is papered with 96 x 48 copies of it rather than one blur.
  const tiled = (fill: string) =>
    buildDocxFromBody(
      `<w:p><w:r><w:drawing>
        <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
          <wp:extent cx="1828800" cy="914400"/>
          <wp:positionH relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>
          <wp:wrapNone/>
          <wp:docPr id="9" name="Papered"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  ${fill}
                </wps:spPr>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing></w:r></w:p>`,
      {
        images: {
          rId20: {
            contentType: 'image/png',
            bytes: buildTinyPng(2, 2, [0, 0, 255, 255]),
            extension: 'png',
          },
        },
      },
    );
  const blip = (inner: string) => `<a:blipFill><a:blip r:embed="rId20"/>${inner}</a:blipFill>`;
  const images = (docx: Uint8Array) =>
    layoutOf(docx).pages[0]!.commands.filter((c) => c.type === 'image');

  it('repeats the picture over the box instead of stretching it', () => {
    const drawn = images(tiled(blip('<a:tile tx="0" ty="0" sx="100000" sy="100000"/>')));
    expect(drawn.length).toBeGreaterThan(100);
    const first = drawn[0] as unknown as { width: number; height: number };
    expect(first.width).toBeCloseTo(1.5, 1);
    expect(first.height).toBeCloseTo(1.5, 1);
  });

  it('…and a stretched one is still the single picture it was', () => {
    const drawn = images(tiled(blip('<a:stretch><a:fillRect/></a:stretch>')));
    expect(drawn.length).toBe(1);
    const only = drawn[0] as unknown as { width: number; height: number };
    expect(only.width).toBeCloseTo(144, 0);
    expect(only.height).toBeCloseTo(72, 0);
  });
});
