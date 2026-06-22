// Table-style resolution (ECMA-376 §17.7.6) — a FlowDoc transform in the
// stage-6 mould: the reader runs it right after parsing, so the tree carries
// final effective cell chrome (shading, borders, margins) and the table-style
// text layer, and no writer has to know table styles exist.
//
// Layer model: a table style contributes its base layer (wholeTable) plus
// conditional region layers (§17.7.6.3) gated by the table's w:tblLook.
// Application order, low → high (§17.7.6.6): wholeTable → column bands → row
// bands → first/last column → first/last row → corner cells; the basedOn
// chain folds root-first within each region.
//
// The style's run/paragraph properties are applied as a FALLBACK under each
// run's/paragraph's direct properties. Strictly, §17.7.6.6 slots the table
// style between docDefaults and the paragraph style; when a cell paragraph
// carries its own styleId AND the table style contests the same field, the
// table style wins here where the spec says the paragraph style should — a
// deliberate v1 simplification (the combination is rare; revisit if a corpus
// document disagrees). Mutates in place — the reader owns the tree
// (resolveBodyStyles sets the precedent).

import type {
  BodyElement,
  CellBorders,
  ParagraphProperties,
  RunProperties,
  Style,
  StyleSheet,
  Table,
  TableCell,
  TableLook,
  TableStyleConditionType,
  TableStyleLayer,
} from '@/core/document-model';

/**
 * Apply table styles across a body (ECMA-376 §17.7.6): for each table, fold its
 * `basedOn` chain and apply the base + `tblLook`-gated conditional region layers
 * to every cell's chrome (shading, borders, margins) and run/paragraph text, as
 * a fallback under direct formatting. Recurses into nested tables and shape
 * text. Mutates the tree IN PLACE and returns it.
 */
export function resolveTableStyles(
  body: ReadonlyArray<BodyElement>,
  sheet: StyleSheet,
): ReadonlyArray<BodyElement> {
  const visit = (el: BodyElement): void => {
    if (el.kind === 'table') {
      applyTableStyle(el.table, sheet);
      for (const row of el.table.rows) {
        for (const cell of row.cells) {
          for (const child of cell.content) visit(child);
        }
      }
    } else if (el.kind === 'shape' && el.shape.text) {
      for (const child of el.shape.text.content) visit(child);
    }
  };
  for (const el of body) visit(el);
  return body;
}

interface FoldedTableStyle {
  readonly base: TableStyleLayer;
  readonly conditions: ReadonlyMap<TableStyleConditionType, TableStyleLayer>;
  readonly rowBandSize: number;
  readonly colBandSize: number;
}

function applyTableStyle(table: Table, sheet: StyleSheet): void {
  const styleId = table.properties.styleId;
  if (!styleId) return;
  const folded = foldTableStyleChain(styleId, sheet);
  if (!folded) return;

  // Table-level chrome: the style's grid/outer borders and default cell
  // margins back-fill the table properties (direct tblPr wins).
  const tp = table.properties as {
    borders?: CellBorders;
    defaultCellMargins?: TableStyleLayer['cellMargins'];
  };
  if (!tp.borders && folded.base.borders) tp.borders = folded.base.borders;
  if (!tp.defaultCellMargins && folded.base.cellMargins) {
    tp.defaultCellMargins = folded.base.cellMargins;
  }

  const look = table.properties.look ?? {};
  const rows = table.rows;
  const colCount = Math.max(
    table.grid.length,
    ...rows.map((r) => r.cells.reduce((s, c) => s + (c.properties.colSpan ?? 1), 0)),
  );

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    let colStart = 0;
    for (const cell of row.cells) {
      const span = cell.properties.colSpan ?? 1;
      const layer = cellLayer(folded, look, {
        firstRow: r === 0,
        lastRow: r === rows.length - 1,
        firstCol: colStart === 0,
        lastCol: colStart + span >= colCount,
        rowBand: bandIndex(r, look.firstRow === true, folded.rowBandSize),
        colBand: bandIndex(colStart, look.firstColumn === true, folded.colBandSize),
        noHBand: look.noHBand === true,
        noVBand: look.noVBand === true,
      });
      colStart += span;
      if (!layer) continue;
      applyLayerToCell(cell, layer);
    }
  }
}

interface MutableCell {
  properties: {
    borders?: CellBorders;
    margins?: TableStyleLayer['cellMargins'];
    shading?: TableStyleLayer['shading'];
  };
  content: ReadonlyArray<BodyElement>;
}

function applyLayerToCell(tableCell: TableCell, layer: TableStyleLayer): void {
  const cell = tableCell as unknown as MutableCell;
  if (layer.shading && !cell.properties.shading) cell.properties.shading = layer.shading;
  if (layer.cellMargins && !cell.properties.margins) cell.properties.margins = layer.cellMargins;
  if (layer.borders) {
    // Per-side fallback: a direct tcBorders side always wins over the style's.
    const own = cell.properties.borders ?? {};
    const merged: CellBorders = { ...layer.borders, ...definedOnly(own) };
    cell.properties.borders = merged;
  }
  if (layer.runProperties || layer.paragraphProperties) {
    for (const el of cell.content) {
      if (el.kind !== 'paragraph') continue; // nested tables resolve themselves
      const p = el.paragraph as unknown as {
        properties: ParagraphProperties;
        runs: ReadonlyArray<{ properties: RunProperties }>;
      };
      if (layer.paragraphProperties) {
        p.properties = { ...layer.paragraphProperties, ...definedOnly(p.properties) };
      }
      if (layer.runProperties) {
        for (const run of p.runs) {
          run.properties = {
            ...layer.runProperties,
            ...definedOnly(run.properties),
          };
        }
      }
    }
  }
}

interface CellPosition {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
  readonly rowBand: number; // 0-based band index, -1 when outside banding
  readonly colBand: number;
  readonly noHBand: boolean;
  readonly noVBand: boolean;
}

// Which band a row/column falls into: rows participating in banding start
// after the first row when the first-row format is on (§17.7.6.6 note); band
// size groups N consecutive rows per band.
function bandIndex(index: number, skipFirst: boolean, bandSize: number): number {
  const adjusted = index - (skipFirst ? 1 : 0);
  if (adjusted < 0) return -1;
  return Math.floor(adjusted / Math.max(1, bandSize));
}

// Merge the applicable region layers, low → high precedence.
function cellLayer(
  folded: FoldedTableStyle,
  look: TableLook,
  pos: CellPosition,
): TableStyleLayer | undefined {
  const layers: Array<TableStyleLayer | undefined> = [folded.base];
  const cond = (t: TableStyleConditionType) => folded.conditions.get(t);

  if (!pos.noVBand && pos.colBand >= 0) {
    layers.push(pos.colBand % 2 === 0 ? cond('band1Vert') : cond('band2Vert'));
  }
  if (!pos.noHBand && pos.rowBand >= 0) {
    layers.push(pos.rowBand % 2 === 0 ? cond('band1Horz') : cond('band2Horz'));
  }
  if (look.firstColumn && pos.firstCol) layers.push(cond('firstCol'));
  if (look.lastColumn && pos.lastCol) layers.push(cond('lastCol'));
  if (look.firstRow && pos.firstRow) layers.push(cond('firstRow'));
  if (look.lastRow && pos.lastRow) layers.push(cond('lastRow'));
  if (look.firstRow && look.firstColumn && pos.firstRow && pos.firstCol) {
    layers.push(cond('nwCell'));
  }
  if (look.firstRow && look.lastColumn && pos.firstRow && pos.lastCol) layers.push(cond('neCell'));
  if (look.lastRow && look.firstColumn && pos.lastRow && pos.firstCol) layers.push(cond('swCell'));
  if (look.lastRow && look.lastColumn && pos.lastRow && pos.lastCol) layers.push(cond('seCell'));

  let out: TableStyleLayer | undefined;
  for (const layer of layers) {
    if (!layer) continue;
    out = out ? mergeLayer(out, layer) : layer;
  }
  return out;
}

// basedOn chain folded root-first: base layers merge into one; conditional
// layers merge per region type.
function foldTableStyleChain(styleId: string, sheet: StyleSheet): FoldedTableStyle | undefined {
  const chain: Array<Style> = [];
  const visited = new Set<string>();
  let cursor: string | undefined = styleId;
  while (cursor && !visited.has(cursor)) {
    visited.add(cursor);
    const style = sheet.styles.get(cursor);
    if (!style || style.type !== 'table') break;
    chain.unshift(style);
    cursor = style.basedOn;
  }
  if (chain.length === 0) return undefined;

  let base: TableStyleLayer = {};
  const conditions = new Map<TableStyleConditionType, TableStyleLayer>();
  let rowBandSize = 1;
  let colBandSize = 1;
  for (const style of chain) {
    // The style's own rPr/pPr are part of its base layer (alongside tblPr).
    const own: TableStyleLayer = {
      ...(style.tableLayer ?? {}),
      ...(Object.keys(style.runProperties).length > 0 && !style.tableLayer?.runProperties
        ? { runProperties: style.runProperties }
        : {}),
      ...(Object.keys(style.paragraphProperties).length > 0 &&
      !style.tableLayer?.paragraphProperties
        ? { paragraphProperties: style.paragraphProperties }
        : {}),
    };
    base = mergeLayer(base, own);
    for (const c of style.tableConditions ?? []) {
      const prev = conditions.get(c.type);
      conditions.set(c.type, prev ? mergeLayer(prev, c.layer) : c.layer);
    }
    if (style.rowBandSize !== undefined) rowBandSize = style.rowBandSize;
    if (style.colBandSize !== undefined) colBandSize = style.colBandSize;
  }
  return { base, conditions, rowBandSize, colBandSize };
}

function mergeLayer(base: TableStyleLayer, override: TableStyleLayer): TableStyleLayer {
  return {
    ...((override.borders ?? base.borders)
      ? { borders: { ...base.borders, ...definedOnly(override.borders ?? {}) } }
      : {}),
    ...((override.cellMargins ?? base.cellMargins)
      ? { cellMargins: { ...base.cellMargins, ...definedOnly(override.cellMargins ?? {}) } }
      : {}),
    ...((override.shading ?? base.shading) ? { shading: override.shading ?? base.shading } : {}),
    ...((override.runProperties ?? base.runProperties)
      ? { runProperties: { ...base.runProperties, ...definedOnly(override.runProperties ?? {}) } }
      : {}),
    ...((override.paragraphProperties ?? base.paragraphProperties)
      ? {
          paragraphProperties: {
            ...base.paragraphProperties,
            ...definedOnly(override.paragraphProperties ?? {}),
          },
        }
      : {}),
  };
}

function definedOnly<T extends object>(obj: T): Partial<T> {
  const out: Partial<T> = {};
  for (const key of Object.keys(obj) as Array<keyof T>) {
    const v = obj[key];
    if (v !== undefined) out[key] = v;
  }
  return out;
}
