// ECMA-376 Part 1 §17.3.1, §20.4 — what a text box's text draws besides its
// words: a table's cell shading (§17.4.33) and rules (§17.4.39), and a
// paragraph's shading (§17.3.1.31). All of it lies ON the box's fill, and Word
// draws the box, then that, then the words (issue #136: a dark-blue cell in a
// white text box came out white).

import { readFileSync } from 'node:fs';

import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { FillItem, PageItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { paintPlan } from '@/layout/page-doc';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const SHD = (fill: string) => `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;

/** A floating text box filled `fill`, 300×100pt, holding `content`. */
const box = (content: string, fill: string): string =>
  `<w:p><w:r><w:drawing>
    <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264"
        behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page"><wp:posOffset>762000</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="page"><wp:posOffset>2540000</wp:posOffset></wp:positionV>
      <wp:extent cx="3810000" cy="1270000"/>
      <wp:wrapNone/>
      <wp:docPr id="1" name="Text Box 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1270000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>
              <a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
            </wps:spPr>
            <wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>
            <wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing></w:r></w:p>`;

const rule = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;

/** The issue's table: a dark-blue cell beside a plain one, ruled all round. */
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(rule).join('') +
  '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>${SHD('25397D')}</w:tcPr>` +
  '<w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Inspection Date</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl><w:p/>';

function commandsOf(body: string): ReadonlyArray<PageItem> {
  const flow = Ream.parse(buildDocxFromBody(body)).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages[0]!.commands;
}

const isBox = (fill: string) => (c: PageItem) =>
  c.type === 'shape' && c.shape.fillColorHex === fill;
const isFill = (fill: string) => (c: PageItem) => c.type === 'fill' && c.fillColorHex === fill;

/** The run of items the page paints in its own order that holds the box. */
function boxRun(commands: ReadonlyArray<PageItem>, fill: string): ReadonlyArray<PageItem> {
  const plan = paintPlan(commands);
  const run = [...plan.ordered, ...plan.pictures].find((r) => r.some(isBox(fill)));
  if (!run) throw new Error('the box paints in no run of its own');
  return run;
}

describe('a table in a text box (issue #136)', () => {
  it('paints its cell shading over the box, and its rules over that', () => {
    const commands = commandsOf(box(TABLE, 'FFFFFF'));
    // Not left in the page's own fill pass, which paints before every shape.
    expect(paintPlan(commands).fills.some(isFill('25397D'))).toBe(false);
    const run = boxRun(commands, 'FFFFFF');
    const boxAt = run.findIndex(isBox('FFFFFF'));
    const cellAt = run.findIndex(isFill('25397D'));
    const ruleAt = run.findIndex((c) => c.type === 'border');
    expect(cellAt).toBeGreaterThan(boxAt);
    expect(ruleAt).toBeGreaterThan(cellAt);
    // Every rule after every fill, so the next row's shading cannot cover one.
    const lastFill = run.reduce((at, c, i) => (c.type === 'fill' ? i : at), -1);
    expect(run.findIndex((c) => c.type === 'border')).toBeGreaterThan(lastFill);
    // The words stay in the text pass, over all of it.
    expect(run.some((c) => c.type === 'line')).toBe(false);
  });

  it('writes the cell after the box in the PDF, and strokes its rules', async () => {
    const pdf = await Ream.parse(buildDocxFromBody(box(TABLE, 'FFFFFF'))).convert('pdf', {
      fonts: FONTS,
    });
    const page = streams(pdf).find((s) => s.includes(' re\n') && s.includes(' rg\n'))!;
    const fillAt = (r: number, g: number, b: number): number => {
      const re = /([\d.]+) ([\d.]+) ([\d.]+) rg/gu;
      for (let m = re.exec(page); m; m = re.exec(page)) {
        const [cr, cg, cb] = [m[1], m[2], m[3]].map(Number) as [number, number, number];
        if (Math.abs(cr - r) < 0.01 && Math.abs(cg - g) < 0.01 && Math.abs(cb - b) < 0.01) {
          return m.index;
        }
      }
      return -1;
    };
    const white = fillAt(1, 1, 1);
    const blue = fillAt(0x25 / 255, 0x39 / 255, 0x7d / 255);
    expect(white).toBeGreaterThanOrEqual(0);
    expect(blue).toBeGreaterThan(white);
    // The grid: black strokes after the cell's fill.
    expect(page.indexOf('0 0 0 RG', blue)).toBeGreaterThan(blue);
  });

  it('marks what it paints with the box in a tagged PDF', async () => {
    // ISO 32000-1 §14.8.2.2 — every painting operator inside marked content.
    const pdf = await Ream.parse(buildDocxFromBody(box(TABLE, 'FFFFFF'))).convert('pdf', {
      fonts: FONTS,
      tagged: true,
    });
    const page = streams(pdf).find((s) => s.includes(' re\n') && s.includes(' rg\n'))!;
    let depth = 0;
    const unmarked: Array<string> = [];
    for (const raw of page.split('\n')) {
      const op = raw.trim();
      if (/\b(BDC|BMC)$/u.test(op)) depth++;
      else if (op === 'EMC') depth--;
      else if (/^(f|F|f\*|S|s|B|B\*|b|b\*)$/u.test(op) && depth === 0) unmarked.push(op);
    }
    expect(unmarked).toEqual([]);
    expect(depth).toBe(0);
  });

  it('writes the cell and the rules after the box in SVG', async () => {
    const svg = new TextDecoder().decode(
      await Ream.parse(buildDocxFromBody(box(TABLE, 'FFFFFF'))).convert('svg', { fonts: FONTS }),
    );
    const boxAt = svg.indexOf('fill="#FFFFFF"');
    const cellAt = svg.indexOf('fill="#25397D"');
    expect(boxAt).toBeGreaterThanOrEqual(0);
    expect(cellAt).toBeGreaterThan(boxAt);
    expect(svg.indexOf('<line', cellAt)).toBeGreaterThan(cellAt);
  });
});

describe('a shaded paragraph (§17.3.1.31)', () => {
  it('paints its band in a text box, over the box’s fill', () => {
    const para =
      `<w:p><w:pPr>${SHD('FFFFFF')}</w:pPr><w:r><w:t>on white</w:t></w:r></w:p>` +
      '<w:p><w:r><w:t>on the box</w:t></w:r></w:p>';
    const run = boxRun(commandsOf(box(para, '25397D')), '25397D');
    const band = run.find(isFill('FFFFFF')) as FillItem | undefined;
    expect(band).toBeDefined();
    expect(run.indexOf(band!)).toBeGreaterThan(run.findIndex(isBox('25397D')));
    // One paragraph's band, not the box's whole height.
    expect(band!.height).toBeLessThan(30);
  });

  it('paints its band in a table cell, over the cell’s fill', () => {
    const cellTable =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>${SHD('25397D')}</w:tcPr>` +
      `<w:p><w:pPr>${SHD('FFFFFF')}</w:pPr><w:r><w:t>on white</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const fills = paintPlan(commandsOf(cellTable + '<w:p/>')).fills;
    const cell = fills.findIndex(isFill('25397D'));
    const band = fills.findIndex(isFill('FFFFFF'));
    expect(cell).toBeGreaterThanOrEqual(0);
    expect(band).toBeGreaterThan(cell);
    // Inside the cell's margins (108 twips a side), as the body's band is.
    expect(fills[band]!.width).toBeCloseTo(fills[cell]!.width - 2 * 5.4, 1);
  });
});

/** Every stream of a PDF, inflated where it is compressed. */
function streams(pdf: Uint8Array): Array<string> {
  const latin1 = new TextDecoder('latin1');
  const text = latin1.decode(pdf);
  const out: Array<string> = [];
  const re = /<<([^]*?)>>\s*stream\r?\n/gu;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    const start = m.index + m[0].length;
    const end = text.indexOf('endstream', start);
    if (!m[1]!.includes('/FlateDecode')) {
      out.push(text.slice(start, end));
      continue;
    }
    try {
      out.push(latin1.decode(unzlibSync(pdf.subarray(start, end))));
    } catch {
      // Not a stream this test reads.
    }
  }
  return out;
}
