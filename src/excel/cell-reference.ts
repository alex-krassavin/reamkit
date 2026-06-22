// ECMA-376 Part 1 §18.18.62 — Cell References.
// Spreadsheet columns are base-26 with digits A-Z (1-indexed: A=1, …, Z=26,
// AA=27, AB=28, …). Rows are plain 1-indexed integers. We expose 0-indexed
// row/col internally — converting at the I/O boundary.

/** A 0-indexed cell address (`A1` → `{ column: 0, row: 0 }`). */
export interface CellAddress {
  readonly column: number;
  readonly row: number;
}

const A = 65;
const Z = 90;
const a = 97;
const z = 122;

/**
 * Parse an A1-style cell reference (§18.18.62) into a 0-indexed {@link CellAddress}.
 * The column letters are base-26 (`A`=1 … `Z`=26, `AA`=27, …); the row is a plain
 * 1-indexed integer. Both are decremented to 0-indexed on the way out.
 *
 * @param ref The reference, e.g. `"AB12"` (case-insensitive column letters).
 * @returns The 0-indexed column/row.
 * @throws Error when `ref` has no column letters or a non-positive-integer row.
 */
export function parseCellRef(ref: string): CellAddress {
  let i = 0;
  let column = 0;
  while (i < ref.length) {
    const code = ref.charCodeAt(i);
    if (code >= A && code <= Z) {
      column = column * 26 + (code - A + 1);
    } else if (code >= a && code <= z) {
      column = column * 26 + (code - a + 1);
    } else {
      break;
    }
    i++;
  }
  if (column === 0) throw new Error(`Invalid cell reference: ${ref}`);
  const rowPart = ref.substring(i);
  const row = Number(rowPart);
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  return { column: column - 1, row: row - 1 };
}

/**
 * Format a 0-indexed {@link CellAddress} back to an A1-style reference — the
 * inverse of {@link parseCellRef} (`{ column: 0, row: 0 }` → `"A1"`).
 */
export function formatCellRef(address: CellAddress): string {
  let n = address.column + 1;
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(A + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return `${s}${address.row + 1}`;
}
