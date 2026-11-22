// What a number format lays out besides text (§18.8.31): the blanks `?` and
// `_x` leave, the spaces it writes, the `*x` fill, and the `#` of a value no
// section admits. Every expectation is what Excel for Mac's own PDF of a probe
// workbook showed (2026-11-20): positions are compared within the page, never
// against Excel's absolute geometry, which is not Windows'.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { FormattedValue } from '@/core/number-format';
import { FontRegistry } from '@/core/font';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import {
  applyNumberFormat,
  applyTextFormat,
  formatNumberParts,
  formatTextParts,
} from '@/core/number-format';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** The parts as a short string: text quoted, a blank as [x], the fill as *x. */
function shape(v: FormattedValue): string {
  const parts = v.parts.map((p) =>
    p.kind === 'text'
      ? JSON.stringify(p.text)
      : p.kind === 'blank'
        ? `[${p.widthOf}]`
        : `*${p.char}`,
  );
  return parts.join(' ') + (v.hashes ? ' #' : '');
}

const num = (code: string, value: number): FormattedValue =>
  formatNumberParts(String(value), 164, new Map([[164, code]]));
const txt = (code: string, value: string): FormattedValue =>
  formatTextParts(value, 164, new Map([[164, code]]));

describe('a number format in the pieces a cell lays it out in', () => {
  it('keeps each blank as wide as the character it stands for', () => {
    expect(shape(num('0.0?', 1.5))).toBe('"1.5" [0]');
    expect(shape(num('_(0_)', 5))).toBe('[(] "5" [)]');
    expect(shape(num('0_W', 5))).toBe('"5" [W]');
    // The blank fraction is as wide as what it would have printed, slash and all.
    expect(shape(num('# ?/?', 3))).toBe('"3 " [0] [/] [0]');
    expect(shape(num('# ??/??', 3.14159))).toBe('"3 " [0] "1/7" [0]');
    expect(shape(num('?,???', 5))).toBe('[0] [,] [0] [0] "5"');
  });

  it('writes the spaces of the format as blanks, which keep their room at an edge', () => {
    expect(shape(num('0" "', 5))).toBe('"5" [ ]');
    expect(shape(num('#,##0.00\\ ', 1234.5))).toBe('"1,234.50" [ ]');
    // Between two words a space is laid out as it is, and stays text.
    expect(shape(num('"Total: "0', 5))).toBe('"Total: 5"');
    expect(shape(num('0"  "0', 15))).toBe('"1" [ ] [ ] "5"');
    expect(shape(num('0" "* ', 5))).toBe('"5" [ ] * ');
  });

  it('marks where the fill goes, once', () => {
    expect(shape(num('"$"* #,##0', 1234))).toBe('"$" *  "1,234"');
    expect(shape(num('_-* #,##0_-;-* #,##0_-', -1234))).toBe('"-" *  "1,234" [-]');
    expect(shape(num('*-0*=', 5))).toBe('*- "5"');
    expect(shape(txt('@*-', 'abc'))).toBe('"abc" *-');
    // The text of a text cell is one piece, its own spaces as they are.
    expect(shape(txt('_(@_)', 'a b'))).toBe('[(] "a b" [)]');
  });

  it('says when no section admits the value, and still writes the number', () => {
    expect(shape(num('"a"0;[<-10]"b"0', -5))).toBe('"-5" #');
    expect(shape(num('"a"0;[<-10]"b"0', 0))).toBe('"0" #');
    expect(shape(num('[<=100]0;[>200]0', 150))).toBe('"150" #');
    expect(num('"a"0;[<-10]"b"0', 5).hashes).toBeUndefined();
    // One section shows such a value as General instead.
    expect(num('[>0]"a"0', -5).hashes).toBeUndefined();
  });

  it('leaves the text of the format as it was: Excel’s string value has no widths', () => {
    expect(applyNumberFormat('1234', 164, new Map([[164, '"$"* #,##0']]))).toBe('$1,234');
    expect(applyNumberFormat('3', 164, new Map([[164, '# ?/?']]))).toBe('3    ');
    expect(applyNumberFormat('3.14159', 164, new Map([[164, '# ??/??']]))).toBe('3  1/7 ');
    expect(applyNumberFormat('-5', 164, new Map([[164, '"a"0;[<-10]"b"0']]))).toBe('-5');
    expect(applyTextFormat('abc', 164, new Map([[164, '@*-']]))).toBe('abc');
    for (const [code, value] of [
      ['?.???', 1.5],
      ['_(0_);(0)', -5],
      ['0" "', 5],
    ] as const) {
      expect(num(code, value).text).toBe(
        applyNumberFormat(String(value), 164, new Map([[164, code]])),
      );
    }
  });
});

/** One token of a laid-out line, where it starts and how wide it is. */
interface PlacedToken {
  readonly text: string;
  readonly x: number;
  readonly width: number;
}

/** The lines of the first page, each as its tokens, top to bottom. */
function lines(xlsx: Uint8Array): Array<{ y: number; tokens: Array<PlacedToken> }> {
  const flow = Ream.parse(xlsx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out: Array<{ y: number; tokens: Array<PlacedToken> }> = [];
  for (const command of laid.pages[0]!.commands) {
    if (command.type !== 'line') continue;
    const c = command as unknown as {
      originX: number;
      baselineY: number;
      line: { tokens: ReadonlyArray<{ text?: string; widthPt: number }> };
    };
    let x = c.originX;
    const tokens: Array<PlacedToken> = [];
    for (const t of c.line.tokens) {
      tokens.push({ text: t.text ?? '', x, width: t.widthPt });
      x += t.widthPt;
    }
    out.push({ y: c.baselineY, tokens });
  }
  return out;
}

/** The token whose text is `text`, on the line that holds it. */
function token(all: ReturnType<typeof lines>, text: string, nth = 0): PlacedToken {
  const hits = all.flatMap((l) => l.tokens.filter((t) => t.text === text));
  const hit = hits[nth];
  if (!hit) throw new Error(`no token "${text}" (#${nth})`);
  return hit;
}

const end = (t: PlacedToken): number => t.x + t.width;

/**
 * A sheet of one column, 16 characters wide: a row per [format, value,
 * alignment], and a plain right-aligned 5 under them to mark where the cell's
 * text ends on the right.
 */
function sheet(
  rows: ReadonlyArray<readonly [string, number | string, (string | undefined)?]>,
): Uint8Array {
  const codes = [...new Set(rows.map((r) => r[0]))];
  const attr = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
  const xfs = rows.map(([code, , align]) => {
    const al = align ? `<alignment horizontal="${align}"/>` : '';
    const id = 164 + codes.indexOf(code);
    return `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"${align ? ' applyAlignment="1"' : ''}>${al}</xf>`;
  });
  const stylesXml =
    `<numFmts count="${codes.length}">` +
    codes.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${attr(c)}"/>`).join('') +
    '</numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    xfs.join('') +
    '</cellXfs>';
  return buildXlsx({
    rows: [...rows.map(([, value], i) => [{ value, styleIndex: i + 1 }]), [{ value: 5 }]],
    columns: [{ min: 1, max: 1, widthChars: 16 }],
    stylesXml,
  });
}

describe('a cell lays out the blanks of its format', () => {
  it('keeps a blank’s room at the right edge, where a space would take none', () => {
    const all = lines(
      sheet([
        ['0_)', 5],
        ['0" "', 5],
        ['0.0?', 1.5],
        ['0.0?', 1.25],
        ['_(0_);(0)', -5],
      ]),
    );
    // The plain right-aligned 5 under them marks where text ends at the edge.
    const edge = end(token(all, '5', 2));
    // Excel's PDF: 5 under `0_)` a parenthesis in from the edge, under `0" "`
    // a space in; (5) against it, its `)` where the blank was; 1.5 and 1.25
    // under `0.0?` with their points one above the other.
    expect(end(token(all, '5', 0))).toBeLessThan(edge - 1);
    expect(end(token(all, '5', 1))).toBeLessThan(edge - 1);
    expect(end(token(all, '5', 1))).toBeGreaterThan(end(token(all, '5', 0)));
    expect(end(token(all, '(5)'))).toBeCloseTo(edge, 3);
    expect(token(all, '1.5').x).toBeCloseTo(token(all, '1.25').x, 3);
    expect(end(token(all, '1.25'))).toBeCloseTo(edge, 3);
  });

  it('holds a fraction that rounds away as wide as the fraction would be', () => {
    const all = lines(
      sheet([
        ['# ?/?', 3],
        ['# ?/?', 3.5],
      ]),
    );
    // "3" of 3 stands where the "3" of 3 1/2 does.
    expect(token(all, '3').x).toBeCloseTo(token(all, '3', 1).x, 3);
  });
});

describe('a cell lays out the fill of its format', () => {
  it('sets the text after the `*` against the right edge, whatever the alignment', () => {
    const rows = [undefined, 'left', 'center', 'right'].map(
      (align) => ['"$"* #,##0', 1234, align] as const,
    );
    const all = lines(sheet(rows));
    const edge = end(token(all, '5'));
    const dollars = [0, 1, 2, 3].map((i) => token(all, '$', i));
    const amounts = [0, 1, 2, 3].map((i) => token(all, '1,234', i));
    // Excel's PDF: $ at the left and 1,234 at the right of a General, left,
    // centre and right-aligned cell alike.
    for (const a of amounts) expect(end(a)).toBeCloseTo(edge, 3);
    // …the $ moved on by no more than the room a space of fill leaves over.
    for (const d of dollars.slice(1)) expect(Math.abs(d.x - dollars[0]!.x)).toBeLessThan(3);
    expect(dollars[0]!.x).toBeLessThan(amounts[0]!.x - 50);
  });

  it('repeats whole characters, the room left over placed by the alignment', () => {
    const all = lines(
      sheet([
        ['*W0', 5],
        ['*W0', 5, 'center'],
        ['*W0', 5, 'right'],
      ]),
    );
    const fills = all
      .map((l) => l.tokens.find((t) => /^W+$/.test(t.text)))
      .filter((t): t is PlacedToken => t !== undefined);
    expect(fills).toHaveLength(3);
    const count = fills[0]!.text.length;
    for (const f of fills) expect(f.text.length).toBe(count);
    const fives = [0, 1, 2].map((i) => token(all, '5', i));
    // The 5 keeps to the right edge in all three.
    for (const f of fives) expect(end(f)).toBeCloseTo(end(fives[0]!), 3);
    // General: the W from the left edge, the room left over between them and
    // the 5 — inside the fill, after its characters. Right: all of that room
    // in front. Centre: half of it each side.
    const spare = fills[0]!.width - fills[2]!.width;
    const lead = (i: number): number => fills[i]!.x - fills[0]!.x;
    expect(spare).toBeGreaterThanOrEqual(0);
    expect(lead(2)).toBeCloseTo(spare, 3);
    expect(lead(1)).toBeCloseTo(spare / 2, 3);
  });

  it('fills the cell with # where the text will not fit beside the fill, text too', () => {
    const all = lines(
      sheet([
        ['"abcdefghijklmnopqrstuvwxyz"* 0', 5],
        ['@*-', 'a long text that runs well past the edge of the cell'],
        ['@*-', 'abc'],
      ]),
    );
    const texts = all.map((l) => l.tokens.map((t) => t.text).join(''));
    expect(texts.filter((t) => /^#+$/.test(t))).toHaveLength(2);
    expect(texts.some((t) => /^abc-+$/.test(t))).toBe(true);
  });
});

describe('a value no section admits', () => {
  it('shows a row of # across the cell, centred, in any alignment', () => {
    const all = lines(
      sheet([
        ['"a"0;[<-10]"b"0', -5],
        ['"a"0;[<-10]"b"0', 0, 'left'],
        ['"a"0;[<-10]"b"0', 5],
      ]),
    );
    const hashes = all
      .map((l) => l.tokens.find((t) => /^#+$/.test(t.text)))
      .filter((t): t is PlacedToken => t !== undefined);
    expect(hashes).toHaveLength(2);
    expect(hashes[1]!.x).toBeCloseTo(hashes[0]!.x, 3);
    expect(hashes[0]!.text).toBe(hashes[1]!.text);
    expect(token(all, 'a5')).toBeDefined();
  });
});

describe('the HTML of a sheet', () => {
  const html = async (xlsx: Uint8Array): Promise<string> =>
    new TextDecoder().decode(await Ream.parse(xlsx).convert('html'));

  it('sizes a blank by its character and lets a fill take the room left', async () => {
    const out = await html(
      sheet([
        ['_(0_)', 5],
        ['"$"* #,##0', 1234],
        ['@*-', 'abc'],
        ['"a"0;[<-10]"b"0', -5],
      ]),
    );
    expect(out).toMatch(/<span class="nf-blank"[^>]* data-w="\)"> <\/span>/);
    expect(out).toMatch(/<p class="nf-line"[^>]*>.*\$.*<span class="nf-fill"[^>]*><\/span>.*1,234/);
    expect(out).toMatch(/<span class="nf-fill"[^>]* data-fill="-{200}"><\/span>/);
    expect(out).toMatch(/<span class="nf-hashes"[^>]* data-fill="#{200}"><\/span>/);
    expect(out).toContain('.nf-blank::before{content:attr(data-w)');
  });

  it('adds nothing to a page whose formats leave no blanks', async () => {
    const out = await html(sheet([['0.00', 5]]));
    expect(out).not.toContain('nf-');
  });
});
