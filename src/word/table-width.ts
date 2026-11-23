// ECMA-376 Part 1 §17.4.88 CT_TblWidth — one reading of a `w:w` / `w:type`
// pair, shared by every element of that type: tblW, tcW, tblInd and the cell
// margins (tblCellMar and tcMar sides).

import { percentSignValue, universalMeasureTwips } from '@/core/po-helpers';

/**
 * What a CT_TblWidth measures once Word has read it.
 *
 * - `pct`: a share of the containing width, in fiftieths of a percent (5000 is
 *   all of it).
 * - `dxa`: an absolute measure in twips.
 * - `auto` / `nil`: no measure of its own. `twips` keeps the bare number the
 *   element carried, for a caller that has always held on to it.
 *
 * `spelled` is true when the value's own spelling (a percent sign or a unit)
 * chose the kind, false when a bare number took it from `type`. A bare number
 * under no type or one outside ST_TblWidth (§17.18.90) reads as `dxa` with
 * `spelled` false, so a caller can still tell the type was never said.
 */
export type TableMeasure = { readonly spelled: boolean } & (
  | { readonly kind: 'pct'; readonly fiftieths: number }
  | { readonly kind: 'dxa'; readonly twips: number }
  | { readonly kind: 'auto' | 'nil'; readonly twips: number }
);

/**
 * §17.4.88 CT_TblWidth — read `w` as Word reads it. `w` is §17.18.107
 * ST_MeasurementOrPercent: a bare number, a percentage with its sign, or a
 * universal measure (§22.9.2.15). A bare number means what `type` says it
 * means. The other two spellings say it themselves, and Word lets them
 * overrule `type`. Measured on #145's probes (Word 16.114):
 *
 * - `"50%"` is half the width whatever the type: `pct`, `dxa`, `auto` and
 *   `nil` all read as "preferred width percent 50".
 * - A fraction of a percent is dropped, not rounded: `"33.6%"` reads 33,
 *   `"99.9%"` 99, `"66.7%"` 66, `"0.5%"` 0. Fiftieths keep their own
 *   precision (`"1666"` reads 33.32%), but not a decimal place (`"1666.6"`
 *   reads the same).
 * - `"2in"` with `type="pct"` is 144pt, not 2880 fiftieths. `"0.5in"` with
 *   `type="auto"` is a 36pt cell margin.
 *
 * @param w The `w` attribute.
 * @param type The `type` attribute.
 * @returns The measure, or undefined when `w` is absent or spelled no way
 *   Word reads.
 */
export function readTableMeasure(
  w: string | undefined,
  type: string | undefined,
): TableMeasure | undefined {
  if (w === undefined) return undefined;
  const percent = percentSignValue(w);
  if (percent !== undefined) {
    return { kind: 'pct', fiftieths: Math.trunc(percent) * 50, spelled: true };
  }
  const n = Number(w);
  if (!Number.isFinite(n)) {
    const twips = universalMeasureTwips(w);
    return twips === undefined ? undefined : { kind: 'dxa', twips, spelled: true };
  }
  if (type === 'pct') return { kind: 'pct', fiftieths: Math.trunc(n), spelled: false };
  if (type === 'auto' || type === 'nil') return { kind: type, twips: n, spelled: false };
  return { kind: 'dxa', twips: n, spelled: false };
}

/**
 * One side of a cell margin (a `w:tblCellMar` or `w:tcMar` child) in twips.
 * The spec gives a margin no `pct` or `auto` measure ("shall be ignored"), and
 * Word reads the two layers apart. On the table's own default (`tblCellMar`)
 * a `pct`, `auto` or `nil` side is ZERO: the text sat flush against the
 * border, for `"10%"` in any type as well. On a cell's own (`tcMar`) such a
 * side is not there at all, and the table's default shows through. A negative
 * inset reads as none (negative-cell-margin-twips.docx; Word put `-720` flush
 * as well).
 *
 * @param w The side's `w` attribute.
 * @param type The side's `type` attribute.
 * @param tableDefault True for a `w:tblCellMar` side, false for `w:tcMar`.
 * @returns Twips, or undefined when the side says nothing a reader keeps.
 */
export function cellMarginTwips(
  w: string | undefined,
  type: string | undefined,
  tableDefault: boolean,
): number | undefined {
  const m = readTableMeasure(w, type);
  if (m === undefined) return undefined;
  if (m.kind === 'dxa') return Math.max(0, m.twips);
  return tableDefault ? 0 : undefined;
}
