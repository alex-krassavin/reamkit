// ECMA-376 Part 1 §18.2.5 / §18.17 — resolving a <definedName> range value
// (the formula syntax used by _xlnm.Print_Area) into a cell range.
//
// A Print_Area value looks like "Sheet1!$A$1:$D$20", may be quoted when the
// sheet name needs it ("'My Sheet'!$A$1:$D$20"), may be a single cell
// ("Sheet1!$A$1"), and may list several comma-separated areas
// ("Sheet1!$A$1:$D$20,Sheet1!$F$1:$H$10"). We resolve to the bounding box of
// every parseable area (a faithful approximation: multiple disjoint print
// areas are rare, and Excel itself prints them in sequence).

import { parseCellRef } from '@/ooxml/spreadsheet/cell-reference';

export interface CellRange {
  readonly startColumn: number;
  readonly startRow: number;
  readonly endColumn: number;
  readonly endRow: number;
}

// Strip the leading "Sheet!" / "'Sheet Name'!" qualifier and the $ absolute
// markers from a single area token, returning just the A1[:A1] part.
function stripQualifier(token: string): string {
  const bang = token.lastIndexOf('!');
  const refPart = bang >= 0 ? token.substring(bang + 1) : token;
  return refPart.replace(/\$/g, '').trim();
}

function parseSingleArea(token: string): CellRange | undefined {
  const ref = stripQualifier(token);
  if (ref.length === 0) return undefined;
  const colon = ref.indexOf(':');
  try {
    if (colon < 0) {
      const a = parseCellRef(ref);
      return { startColumn: a.column, startRow: a.row, endColumn: a.column, endRow: a.row };
    }
    const a = parseCellRef(ref.substring(0, colon));
    const b = parseCellRef(ref.substring(colon + 1));
    return {
      startColumn: Math.min(a.column, b.column),
      startRow: Math.min(a.row, b.row),
      endColumn: Math.max(a.column, b.column),
      endRow: Math.max(a.row, b.row),
    };
  } catch {
    // Whole-row ("$1:$1") / whole-column ("$A:$A") refs, or anything malformed,
    // carry no usable bounding box for a print AREA — skip the token.
    return undefined;
  }
}

// _xlnm.Print_Titles names repeated print rows/columns, e.g. "Sheet1!$1:$2"
// (rows 1-2) or "Sheet1!$A:$B,Sheet1!$1:$1" (cols A-B + row 1). We extract the
// ROW range only — column titles repeat across horizontal page breaks, which we
// don't paginate. Returns 0-indexed inclusive rows, or undefined if none.
export function parseTitleRowRange(
  value: string,
): { readonly startRow: number; readonly endRow: number } | undefined {
  if (!value) return undefined;
  for (const token of value.split(',')) {
    const ref = stripQualifier(token); // "$1:$2" → "1:2"; "$A:$B" → "A:B"
    const m = /^(\d+)(?::(\d+))?$/.exec(ref);
    if (!m) continue; // not a pure row range (a column ref or malformed)
    const a = Number(m[1]);
    const b = m[2] !== undefined ? Number(m[2]) : a;
    if (!Number.isInteger(a) || a < 1 || !Number.isInteger(b) || b < 1) continue;
    return { startRow: Math.min(a, b) - 1, endRow: Math.max(a, b) - 1 };
  }
  return undefined;
}

export function parseAreaRef(value: string): CellRange | undefined {
  if (!value) return undefined;
  let box: CellRange | undefined;
  for (const token of value.split(',')) {
    const area = parseSingleArea(token);
    if (!area) continue;
    box = box
      ? {
          startColumn: Math.min(box.startColumn, area.startColumn),
          startRow: Math.min(box.startRow, area.startRow),
          endColumn: Math.max(box.endColumn, area.endColumn),
          endRow: Math.max(box.endRow, area.endRow),
        }
      : area;
  }
  return box;
}
