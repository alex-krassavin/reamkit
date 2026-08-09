// §20.1.7.6 `a:xfrm rot` — a shape that turns turns its WORDS with it. The
// outline always took the turn; the text was laid out flat and drawn flat, so a
// label set on its side came back lying across whatever it crossed.
// 160F-2019.pdf sets "Nature" down the middle of a column that way.

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

/** A 144×72pt text box holding one word, turned by the `a:xfrm` attributes. */
const turnedBox = (rot: string, bodyPr = '<wps:bodyPr/>'): string =>
  `<w:p><w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Turned"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:spPr><a:xfrm${rot}><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
            </wps:spPr>
            <wps:txbx><w:txbxContent><w:p><w:r><w:t>Nature</w:t></w:r></w:p></w:txbxContent></wps:txbx>
            ${bodyPr}
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;

interface LineCommand {
  readonly originX: number;
  readonly baselineY: number;
  readonly rotationDeg?: number;
  readonly line: { tokens: ReadonlyArray<{ text?: string }> };
}

function wordLine(rot: string, bodyPr?: string): LineCommand {
  const flow = Ream.parse(buildDocxFromBody(turnedBox(rot, bodyPr))).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const found = laid.pages[0]!.commands.filter((c) => c.type === 'line')
    .map((c) => c as unknown as LineCommand)
    .find((c) => c.line.tokens.map((t) => t.text ?? '').join('') === 'Nature');
  if (!found) throw new Error('the shape drew no words');
  return found;
}

describe('a shape’s text turns with the shape (§20.1.7.6)', () => {
  it('leaves an upright box’s words upright', () => {
    expect(wordLine('').rotationDeg).toBeUndefined();
  });

  it('turns the words a quarter turn when the box turns one', () => {
    // `rot` is clockwise; the page frame runs y-up, where positive turns the
    // other way, so a quarter turn clockwise is −90°.
    expect(wordLine(' rot="5400000"').rotationDeg).toBeCloseTo(-90, 5);
  });

  it('adds the box’s turn to a quarter turn of its own', () => {
    // §20.1.10.83 `vert` reads top-to-bottom (a quarter turn clockwise, −90°
    // here); a box turned a half turn on top of it reads the other way up.
    // shape-text-rotate.pptx is exactly that pair, and each half alone is wrong.
    expect(wordLine('', '<wps:bodyPr vert="vert"/>').rotationDeg).toBeCloseTo(-90, 5);
    expect(wordLine(' rot="10800000"', '<wps:bodyPr vert="vert"/>').rotationDeg).toBeCloseTo(
      -270,
      5,
    );
  });

  it('leaves the words level when the box asks to be upright', () => {
    // §20.1.10.55 — bnc762542.xlsx turns each legend label a quarter and asks
    // for this, and every reader draws those labels lying flat.
    const level = wordLine(' rot="5400000"', '<wps:bodyPr upright="1"/>');
    expect(level.rotationDeg).toBeUndefined();
    // Level and in the box's own upright frame, exactly where no turn puts it.
    expect(level.originX).toBeCloseTo(wordLine('').originX, 5);
    expect(level.baselineY).toBeCloseTo(wordLine('').baselineY, 5);
  });

  it('leaves the words level when the box is mirrored', () => {
    // §20.1.7.6 — a flip mirrors the OUTLINE, and mirrored words are not what
    // any reader draws: rot180-flipv.docx and rot270-flipv.docx turn a triangle
    // and flip it back, and the reference sets its label flat across both.
    expect(wordLine(' rot="10800000" flipV="1"').rotationDeg).toBeUndefined();
    expect(wordLine(' rot="16200000" flipV="1"').rotationDeg).toBeUndefined();
    expect(wordLine(' flipH="1"').rotationDeg).toBeUndefined();
  });

  it('spins the line about the box’s own centre, not about its corner', () => {
    // A half turn maps every point to its opposite through the centre, so the
    // two origins straddle it: their midpoints are the centre itself.
    const upright = wordLine('');
    const turned = wordLine(' rot="10800000"');
    expect(turned.rotationDeg).toBeCloseTo(-180, 5);
    // The box is 144×72pt, so a half turn moves the origin by twice its
    // distance from the centre in each axis — never leaving the box.
    expect(Math.abs(turned.originX - upright.originX)).toBeLessThan(144);
    expect(Math.abs(turned.baselineY - upright.baselineY)).toBeLessThan(72);
    expect(turned.originX).not.toBeCloseTo(upright.originX, 1);
  });
});
