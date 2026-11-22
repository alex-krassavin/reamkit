// ECMA-376 Part 1 §17.3.2.6 — a run's colour may be the AUTOMATIC one, which
// the standard leaves "to the consumer to automatically choose an appropriate
// color based on the background color behind the run's content". Word chooses
// black, or white over a dark background, and a document relies on it: issue
// #137's header cell is filled 25397D and its words name no colour, and drawn
// black they could not be read.
//
// Word's rule, measured on our own probes in Word for Mac 16.115 (its PDF
// export, text colour per word read back with `mutool trace`):
//   - the background is judged by its Rec. 601 luma, 0.299 R + 0.587 G +
//     0.114 B: of 182 cell fills — grey, red, green and blue ramps, then each
//     step around the switch — not one falls on the wrong side of it, where
//     the best threshold still misjudges 5 under the weights 0.30/0.59/0.11
//     (and LibreOffice's 76/151/29), 40 under Rec. 709, 41 under linear
//     luminance and 60 under HSL lightness;
//   - the text turns white below 59.77 on the 0–255 scale (the last white at
//     59.765, 661980; the first black at 59.776, 681880) — 0x3C00 on a 16-bit
//     scale, 15360/257, which no 8-bit colour lands on exactly;
//   - the background is the run's own shading, else its paragraph's, else
//     its table cell's, else its text box's fill, each as painted (a pattern
//     blended in: `solid` 25397D over white turns it white, `pct50` of black
//     does not);
//   - a highlight does not count: auto text stays black on every one of the
//     sixteen, the black one included, and on a yellow highlight in a dark
//     cell it turns white for the cell;
//   - a stated 000000 stays black on any background; a stated `auto` is the
//     automatic colour.
// The page's own colour (§17.2.1 `w:background`) is left out: Word's PDF does
// not print it, and keeps the text black.
//
// A FlowDoc transform for the targets that paint a page (PDF, SVG, HTML):
// it states the colour Word would choose, on a copy, and leaves the document
// a writer re-serializes saying `auto`.

import type { BodyElement, Paragraph, ShapeBlock, StyleSheet, Table } from '@/core/document-model';

import { resolveParagraphProperties, resolveRunProperties } from '@/core/style-cascade/resolver';

/** Rec. 601 luma × 1000 × 257 below which Word sets automatic text white: 0x3C00 × 1000. */
const DARK_BELOW = 15_360_000;

/**
 * Whether Word sets automatic text white over this background (§17.3.2.6):
 * its Rec. 601 luma is below 0x3C00 on a 16-bit scale (59.77 of 255).
 *
 * @param hex The background, as 6 hex digits.
 * @returns `true` for a background dark enough to turn automatic text white.
 */
export function isDarkBehindAutoText(hex: string): boolean {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  if (!Number.isFinite(r + g + b)) return false;
  return (299 * r + 587 * g + 114 * b) * 257 < DARK_BELOW;
}

/**
 * Word's automatic text colour made explicit: every run whose colour is the
 * automatic one (§17.3.2.6) and that stands on a dark background — its own
 * shading, its paragraph's, its cell's or its text box's fill — is given white.
 * Recurses into tables and shape text. Returns a new tree where anything
 * changed and the same objects where nothing did; the input is not touched.
 *
 * @param body     The body, its properties already resolved.
 * @param sheet    The style sheet runs resolve against (the empty one for a
 *                 resolved FlowDoc).
 * @param behindHex The background the body stands on, if it is not the page.
 * @returns The body with those runs' colour stated.
 */
export function resolveAutoTextColors(
  body: ReadonlyArray<BodyElement>,
  sheet: StyleSheet,
  behindHex?: string,
): ReadonlyArray<BodyElement> {
  return mapShared(body, (el) => visitElement(el, sheet, behindHex));
}

/**
 * {@link resolveAutoTextColors} over each body of a set — a document's notes,
 * or its header and footer bands — each on the page.
 *
 * @param bodies The bodies, by id.
 * @param sheet  The style sheet runs resolve against.
 * @returns The same ids, each body with its runs' automatic colour stated.
 */
export function resolveAutoTextColorsOfEach(
  bodies: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  sheet: StyleSheet,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  return new Map(
    [...bodies].map(([id, body]) => [id, resolveAutoTextColors(body, sheet)] as const),
  );
}

/** `items` mapped by `f` — or `items` itself, when `f` gave every item back unchanged. */
function mapShared<T>(items: ReadonlyArray<T>, f: (item: T) => T): ReadonlyArray<T> {
  let out: Array<T> | undefined;
  for (const [i, item] of items.entries()) {
    const next = f(item);
    if (next !== item) out ??= items.slice(0, i);
    out?.push(next);
  }
  return out ?? items;
}

function visitElement(el: BodyElement, sheet: StyleSheet, behind: string | undefined): BodyElement {
  if (el.kind === 'paragraph') {
    const paragraph = visitParagraph(el.paragraph, sheet, behind);
    return paragraph === el.paragraph ? el : { ...el, paragraph };
  }
  if (el.kind === 'table') {
    const table = visitTable(el.table, sheet, behind);
    return table === el.table ? el : { ...el, table };
  }
  if (el.kind === 'shape') {
    const shape = visitShape(el.shape, sheet, behind);
    return shape === el.shape ? el : { ...el, shape };
  }
  return el;
}

function visitParagraph(p: Paragraph, sheet: StyleSheet, behind: string | undefined): Paragraph {
  const paragraphBehind =
    resolveParagraphProperties(p.properties, sheet).shading?.colorHex ?? behind;
  const runs = mapShared(p.runs, (run) => {
    const resolved = resolveRunProperties(run.properties, p.properties, sheet);
    if (resolved.colorAuto !== true) return run;
    // The run's own shading, never its highlight, even where that shows.
    const under = resolved.shadingColorHex ?? paragraphBehind;
    if (under === undefined || !isDarkBehindAutoText(under)) return run;
    const { colorAuto: _auto, ...stated } = run.properties;
    return { ...run, properties: { ...stated, colorHex: 'FFFFFF' } };
  });
  return runs === p.runs ? p : { ...p, runs };
}

function visitTable(table: Table, sheet: StyleSheet, behind: string | undefined): Table {
  const rows = mapShared(table.rows, (row) => {
    const cells = mapShared(row.cells, (cell) => {
      const under = cell.properties.shading?.colorHex ?? behind;
      const content = resolveAutoTextColors(cell.content, sheet, under);
      return content === cell.content ? cell : { ...cell, content };
    });
    return cells === row.cells ? row : { ...row, cells };
  });
  return rows === table.rows ? table : { ...table, rows };
}

function visitShape(shape: ShapeBlock, sheet: StyleSheet, behind: string | undefined): ShapeBlock {
  // A box with no fill of its own shows what is behind it; one filled with a
  // gradient or a picture has no one colour to judge by.
  const fill = shape.fill.kind === 'solid' ? shape.fill.colorHex : undefined;
  const under = fill ?? (shape.fill.kind === 'none' ? behind : undefined);
  const text = shape.text;
  const content = text ? resolveAutoTextColors(text.content, sheet, under) : undefined;
  const children = shape.children
    ? mapShared(shape.children, (child) => {
        const next = visitShape(child.shape, sheet, under);
        return next === child.shape ? child : { ...child, shape: next };
      })
    : undefined;
  const newText = text && content && content !== text.content ? { ...text, content } : undefined;
  const newChildren = children !== shape.children ? children : undefined;
  if (!newText && !newChildren) return shape;
  return {
    ...shape,
    ...(newText ? { text: newText } : {}),
    ...(newChildren ? { children: newChildren } : {}),
  };
}
