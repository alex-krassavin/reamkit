// §20.1.4.2.24/§20.1.4.2.25 — the style a slide table wears, and the theme
// slots its background points at. A table states almost nothing itself: the
// colours come from `ppt/tableStyles.xml`, and that part in turn points into the
// theme's fill styles by INDEX.

import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import type { PoNode } from '@/core/po-helpers';
import { makeColorResolver } from '@/core/drawingml/colors';
import { parseTheme, parseThemeFillStyles } from '@/core/drawingml/theme-parser';
import { cellStyle } from '@/pptx/table-style';
import { poIs } from '@/core/po-helpers';

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
});

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

// The Office theme's shape: a solid slot, then two gradients of it.
function themeXml(pretty: boolean): Uint8Array {
  const gap = pretty ? '\n      ' : '';
  return new TextEncoder().encode(
    `<a:theme xmlns:a="${A}"><a:themeElements>` +
      `<a:clrScheme name="t"><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1></a:clrScheme>` +
      `<a:fmtScheme><a:fillStyleLst>${gap}` +
      `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>${gap}` +
      `<a:gradFill><a:gsLst>` +
      `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>` +
      `<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/></a:schemeClr></a:gs>` +
      `</a:gsLst></a:gradFill>${gap}` +
      `<a:gradFill><a:gsLst>` +
      `<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/></a:schemeClr></a:gs>` +
      `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:gs>` +
      `</a:gsLst></a:gradFill>${gap}` +
      `</a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
  );
}

const NO_FLAGS = {
  firstRow: false,
  lastRow: false,
  firstCol: false,
  lastCol: false,
  bandRow: false,
  bandCol: false,
} as const;

function styleNode(inner: string): PoNode {
  const tree = parser.parse(
    `<a:tblStyle xmlns:a="${A}" styleId="s">${inner}</a:tblStyle>`,
  ) as Array<PoNode>;
  const found = tree.find((n) => poIs(n, 'a:tblStyle'));
  if (!found) throw new Error('no style');
  return found;
}

describe('slide table: the background a style points at', () => {
  it('counts a theme slot by its place among the FILLS, not among the nodes', () => {
    // A theme written one element per line puts a whitespace node between each
    // pair of them. Counted in, `a:fillRef idx="3"` reached the run of spaces
    // after the first gradient and the table came out the flat accent.
    expect(parseThemeFillStyles(themeXml(false))).toHaveLength(3);
    expect(parseThemeFillStyles(themeXml(true))).toHaveLength(3);
  });

  it('takes the slot a fillRef names, with the transforms that slot states', () => {
    const fills = parseThemeFillStyles(themeXml(true));
    const colors = makeColorResolver(parseTheme(themeXml(true)));
    const shading = (idx: string): string | undefined =>
      cellStyle(
        styleNode(
          `<a:tblBg><a:fillRef idx="${idx}"><a:schemeClr val="accent1"/></a:fillRef></a:tblBg>` +
            `<a:wholeTbl><a:tcStyle><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>`,
        ),
        NO_FLAGS,
        { row: 0, rowCount: 2, col: 0, colCount: 1 },
        colors,
        { fills },
      ).shadingHex;
    // Slot 1 is the accent itself; slot 2 is a pale tint of it and slot 3 a
    // shade. Read as the bare reference colour, all three came out the accent.
    expect(shading('1')).toBe('5B9BD5');
    const tinted = Number.parseInt(shading('2') ?? '0', 16);
    const shaded = Number.parseInt(shading('3') ?? '0', 16);
    expect(tinted).toBeGreaterThan(0x5b9bd5);
    expect(shaded).toBeLessThan(0x5b9bd5);
  });

  it('lays a part fill of less than full alpha over the background', () => {
    const fills = parseThemeFillStyles(themeXml(true));
    const colors = makeColorResolver(parseTheme(themeXml(true)));
    const style = styleNode(
      `<a:tblBg><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></a:tblBg>` +
        `<a:band1H><a:tcStyle><a:fill><a:solidFill>` +
        `<a:schemeClr val="lt1"><a:alpha val="20000"/></a:schemeClr>` +
        `</a:solidFill></a:fill></a:tcStyle></a:band1H>`,
    );
    const flags = { ...NO_FLAGS, bandRow: true };
    const row = (i: number): string | undefined =>
      cellStyle(style, flags, { row: i, rowCount: 2, col: 0, colCount: 1 }, colors, { fills })
        .shadingHex;
    // A fifth of white over the ACCENT, not a fifth of white over the page:
    // flattened against white the band came out white and the table striped.
    expect(row(1)).toBe('5B9BD5');
    expect(row(0)).toBe('7CAFDD');
  });
});
