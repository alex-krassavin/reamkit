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
// The style's run/paragraph properties are carried on the cell's paragraphs as
// the layer §17.7.6.6 makes them: between docDefaults and the paragraph's own
// style, which the cascade merges in at that rank. Written onto the paragraph
// and its runs instead — as if the document had stated them there — the table
// style outranked every style a cell's text names, and tdf119054.docx's
// headings lost the 18pt their style leaves after each of them. Mutates in
// place — the reader owns the tree (resolveBodyStyles sets the precedent).

import type {
  BodyElement,
  CellBorders,
  CellMargins,
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

import type { Pt } from '@/core/ir';
import { twipsToPt } from '@/core/ir';
import { mergeRunPartial } from '@/core/style-cascade/resolver';

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
    indentPt?: TableStyleLayer['indentPt'];
  };
  // §17.7.6 — the table's OWN `w:tblBorders` outrank the style's conditional
  // regions too, so which sides it states has to be known before they are
  // back-filled from the style below.
  const directBorders = tp.borders;
  // §17.4.39 / §17.4.23 — and they are back-filled SIDE BY SIDE: an edge the
  // table leaves out "shall have the border specified by the associated table
  // style". Taken as a whole set, a table that turned off its frame and its
  // verticals over TableGrid lost the style's rules between its rows too
  // (issue #135), where Word draws them; so does a table that only recolours
  // its top edge, which kept that edge and nothing else.
  if (folded.base.borders) {
    tp.borders = { ...folded.base.borders, ...definedOnly(tp.borders ?? {}) };
  }
  // §17.4.43 — the table's OWN `w:tblCellMar` beats its style's SIDE BY SIDE,
  // and the style's sides fill the rest: under TableGrid (108 each way) a
  // direct `left` of 720 alone insets the text 36pt from the left and 5.4pt
  // from the right in Word 16.114 (style-cellmar probes t02-t04, t18).
  const directMargins = tp.defaultCellMargins;
  if (folded.base.cellMargins) {
    tp.defaultCellMargins = { ...folded.base.cellMargins, ...definedOnly(directMargins ?? {}) };
  }
  if (tp.indentPt === undefined && folded.base.indentPt !== undefined) {
    tp.indentPt = folded.base.indentPt;
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
    const placed: Array<{ cell: TableCell; pos: CellPosition }> = [];
    for (const cell of row.cells) {
      const span = cell.properties.colSpan ?? 1;
      placed.push({
        cell,
        pos: {
          // §17.4.7 — a row may CLAIM the first/last-row format whatever its
          // index is, and Word's calendar templates do exactly that for their
          // second header row.
          firstRow: r === 0 || row.properties.conditional?.firstRow === true,
          lastRow: r === rows.length - 1 || row.properties.conditional?.lastRow === true,
          firstCol: colStart === 0,
          lastCol: colStart + span >= colCount,
          rowBand: bandIndex(r, look.firstRow === true, folded.rowBandSize),
          colBand: bandIndex(colStart, look.firstColumn === true, folded.colBandSize),
          noHBand: look.noHBand === true,
          noVBand: look.noVBand === true,
        },
      });
      colStart += span;
    }
    const rowMargins = rowRegionMargins(placed.map((p) => regionLayers(folded, look, p.pos)));
    for (const { cell, pos } of placed) {
      resolveCellMargins(cell, rowMargins, folded.base.cellLevelMargins, directMargins);
      const layer = cellLayer(folded, look, pos);
      if (!layer) continue;
      applyLayerToCell(cell, layer, directBorders, pos);
    }
  }
}

const SIDES = ['top', 'left', 'bottom', 'right'] as const;
type MarginSides = Partial<Record<(typeof SIDES)[number], Pt>>;

/**
 * The margins a table style's REGIONS give one row. Word lets them reach the
 * whole row whichever of its cells the region names: a firstCol, lastCol,
 * nwCell or band1Vert margin of 1080 insets every cell of the rows it touches
 * by 54pt, while the same regions' shading stays in their own cells
 * (style-cellmar probes c01-c06, x05). They are laid down cell by cell, left
 * to right, each cell's regions in their §17.7.6 order: a firstRow left of
 * 1080 over a nwCell left of 0 comes out 54pt, as the second cell's firstRow
 * has the last word (q06), and a lastCol beats a firstCol (q04).
 */
function rowRegionMargins(cells: ReadonlyArray<ReadonlyArray<TableStyleLayer>>): {
  readonly table: MarginSides;
  readonly cell: MarginSides;
} {
  const table: MarginSides = {};
  const cell: MarginSides = {};
  for (const layers of cells) {
    for (const layer of layers) {
      Object.assign(table, definedOnly(layer.cellMargins ?? {}));
      Object.assign(cell, definedOnly(layer.cellLevelMargins ?? {}));
    }
  }
  return { table, cell };
}

/**
 * Settle one cell's margins, side by side, highest first (Word 16.114,
 * style-cellmar probes): the cell's own `w:tcMar`; the cell-level `w:tcMar`
 * of the style's regions, then of the style as a whole, both above the
 * table's own `w:tblCellMar` (t16, t20); the table-level `w:tblCellMar` of
 * the style's regions, but under the table's own (t15, x06, x08, q08); and
 * last the table's default, which the layout reads from the table. A region's
 * side is written into the cell only where it outranks that default.
 */
function resolveCellMargins(
  tableCell: TableCell,
  row: ReturnType<typeof rowRegionMargins>,
  styleCellLevel: CellMargins | undefined,
  direct: CellMargins | undefined,
): void {
  const cell = tableCell as unknown as MutableCell;
  const fromRegions: MarginSides = {};
  for (const side of SIDES) {
    const v = row.table[side];
    if (v !== undefined && direct?.[side] === undefined) fromRegions[side] = v;
  }
  const merged: MarginSides = {
    ...fromRegions,
    ...definedOnly(styleCellLevel ?? {}),
    ...row.cell,
    ...definedOnly(cell.properties.margins ?? {}),
  };
  if (Object.keys(merged).length > 0) cell.properties.margins = merged;
}

interface MutableCell {
  properties: {
    borders?: CellBorders;
    margins?: TableStyleLayer['cellMargins'];
    shading?: TableStyleLayer['shading'];
  };
  content: ReadonlyArray<BodyElement>;
}

function applyLayerToCell(
  tableCell: TableCell,
  layer: TableStyleLayer,
  // §17.7.6 — the sides the table states in its OWN `w:tblBorders`, which the
  // style's conditional regions do not get to overrule, and where this cell
  // sits (which decides whether an edge answers to `top`/`bottom`/`left`/
  // `right` or to `insideH`/`insideV`).
  directBorders: CellBorders | undefined,
  pos: { firstRow: boolean; lastRow: boolean; firstCol: boolean; lastCol: boolean },
): void {
  const cell = tableCell as unknown as MutableCell;
  if (layer.shading && !cell.properties.shading) cell.properties.shading = layer.shading;
  // The layer's margins are settled by resolveCellMargins: written in here,
  // the style's whole-table `w:tblCellMar` became every cell's own and beat
  // the table's `w:tblCellMar`, which Word ranks above it.
  if (layer.borders) {
    // Per-side fallback: a direct tcBorders side always wins over the style's.
    const own = cell.properties.borders ?? {};
    // …and so does a side the TABLE states. table-style-border.docx rules a
    // black grid in its `w:tblPr` over a style whose corner regions spell the
    // same edges away, and taking the region's word for it left the table with
    // no right edge, no bottom and no rule between its two rows.
    const fromStyle = { ...layer.borders };
    if (directBorders) {
      const stated = (side: 'top' | 'bottom' | 'left' | 'right', outer: boolean): boolean =>
        (outer
          ? directBorders[side]
          : directBorders[side === 'top' || side === 'bottom' ? 'insideH' : 'insideV']) !==
        undefined;
      if (stated('top', pos.firstRow)) delete fromStyle.top;
      if (stated('bottom', pos.lastRow)) delete fromStyle.bottom;
      if (stated('left', pos.firstCol)) delete fromStyle.left;
      if (stated('right', pos.lastCol)) delete fromStyle.right;
    }
    const merged: CellBorders = { ...fromStyle, ...definedOnly(own) };
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
        // §17.7.2 — carried apart from the direct properties, because the
        // cascade ranks it UNDER the paragraph's own style; merged in as if
        // the paragraph stated it, a table style's `w:spacing` beat the
        // heading style of the cell's text (tdf119054, tdf118947_tableStyle).
        p.properties = {
          ...p.properties,
          tableStyle: { ...p.properties.tableStyle, ...layer.paragraphProperties },
        };
      }
      if (layer.runProperties) {
        // §17.7.2 again, on the run side: the table style ranks under the
        // paragraph's style AND under the character style a run names, so it
        // is carried on the paragraph rather than pushed onto each run.
        // Pushed on, tdf118812_tableStyles-comprehensive.docx's every cell
        // came out in the table style's 8pt blue where both references keep
        // the colour and size of the style each paragraph is written in.
        // The paragraph MARK reads the same carrier: a cell with no runs still
        // has one, and its size is the row's height (conditionalstyles-
        // tbllook.docx sets its first column in 36pt).
        p.properties = {
          ...p.properties,
          inheritedRun: mergeRunPartial(p.properties.inheritedRun ?? {}, layer.runProperties),
        };
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
// size groups N consecutive rows per band, and a size of 0 bands nothing.
function bandIndex(index: number, skipFirst: boolean, bandSize: number): number {
  const adjusted = index - (skipFirst ? 1 : 0);
  if (adjusted < 0 || bandSize <= 0) return -1;
  return Math.floor(adjusted / bandSize);
}

// Merge the applicable region layers, low → high precedence.
function cellLayer(
  folded: FoldedTableStyle,
  look: TableLook,
  pos: CellPosition,
): TableStyleLayer | undefined {
  // §17.7.6 — the style's WHOLE-TABLE borders are a table-level default, and
  // the table's own `w:tblBorders` beat them (applyTableStyle back-fills them
  // onto the table when it declares none). Pushed down onto every cell they
  // beat the table's instead: all_gaps_word.docx spells its TableGrid away
  // edge by edge and we boxed every one of its 54 cells.
  const { borders: _tableWide, ...withoutBorders } = folded.base;
  const base: TableStyleLayer = withoutBorders;
  let out: TableStyleLayer | undefined;
  for (const layer of [base, ...regionLayers(folded, look, pos)]) {
    out = out ? mergeLayer(out, layer) : layer;
  }
  return out;
}

// The conditional regions that apply to a cell, low → high precedence.
function regionLayers(
  folded: FoldedTableStyle,
  look: TableLook,
  pos: CellPosition,
): Array<TableStyleLayer> {
  const layers: Array<TableStyleLayer | undefined> = [];
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
  return layers.filter((l): l is TableStyleLayer => l !== undefined);
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

  // A chain that never reaches Word's built-in Normal Table starts from no
  // margins at all: a table style based on nothing that states none insets
  // its text by 0pt in Word, not by the 5.4pt it would inherit (style-cellmar
  // probes t06, t07, v02, v03).
  let base: TableStyleLayer = chain.some(isNormalTable) ? {} : { cellMargins: NO_MARGINS };
  const conditions = new Map<TableStyleConditionType, TableStyleLayer>();
  // §17.7.6.7 / §17.7.6.5 — a style that states no band size, here or up its
  // chain, bands NOTHING in Word 16.114: neither the shading nor the margins
  // of its band1Horz and band1Vert regions show, and a size of 0 does the
  // same. A size stated by a basedOn parent counts (style-cellmar probes
  // scm14 n01-n06, t21).
  let rowBandSize = 0;
  let colBandSize = 0;
  for (const style of chain) {
    // The style's own rPr/pPr are part of its base layer (alongside tblPr).
    const own: TableStyleLayer = {
      ...withNormalTableMargins(style),
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
      const kept = prev && c.layer.restatesMargins ? withoutMargins(prev) : prev;
      conditions.set(c.type, kept ? mergeLayer(kept, c.layer) : c.layer);
    }
    if (style.rowBandSize !== undefined) rowBandSize = style.rowBandSize;
    if (style.colBandSize !== undefined) colBandSize = style.colBandSize;
  }
  return { base, conditions, rowBandSize, colBandSize };
}

// §17.7.4.9 — Word knows its built-in styles by name, and its "Normal Table"
// insets cells by 108 twips left and right and 0 top and bottom whatever the
// document's own definition says: one that states 720 each way still insets
// by 5.4pt, under any table style based on it and in compatibility modes 12,
// 14 and 15 alike, while the same definition under another name insets by
// 36pt (style-cellmar probes scm2, scm6, scm7, scm12).
const NORMAL_TABLE_MARGINS: CellMargins = {
  top: twipsToPt(0),
  left: twipsToPt(108),
  bottom: twipsToPt(0),
  right: twipsToPt(108),
};
const NO_MARGINS: CellMargins = {
  top: twipsToPt(0),
  left: twipsToPt(0),
  bottom: twipsToPt(0),
  right: twipsToPt(0),
};

function isNormalTable(style: Style): boolean {
  return style.name?.toLowerCase() === 'normal table';
}

function withNormalTableMargins(style: Style): TableStyleLayer {
  const layer = style.tableLayer ?? {};
  if (!isNormalTable(style)) return layer;
  return { ...withoutMargins(layer), cellMargins: NORMAL_TABLE_MARGINS };
}

function withoutMargins(layer: TableStyleLayer): TableStyleLayer {
  const { cellMargins: _table, cellLevelMargins: _cell, ...rest } = layer;
  return rest;
}

function mergeLayer(base: TableStyleLayer, override: TableStyleLayer): TableStyleLayer {
  return {
    ...((override.borders ?? base.borders)
      ? { borders: { ...base.borders, ...definedOnly(override.borders ?? {}) } }
      : {}),
    ...((override.cellMargins ?? base.cellMargins)
      ? { cellMargins: { ...base.cellMargins, ...definedOnly(override.cellMargins ?? {}) } }
      : {}),
    // Side by side down the chain, as the table-level margins are: a derived
    // style's `w:tcMar` left of 0 keeps its parent's right of 1080 (k02).
    ...((override.cellLevelMargins ?? base.cellLevelMargins)
      ? {
          cellLevelMargins: {
            ...base.cellLevelMargins,
            ...definedOnly(override.cellLevelMargins ?? {}),
          },
        }
      : {}),
    ...((override.indentPt ?? base.indentPt) !== undefined
      ? { indentPt: override.indentPt ?? base.indentPt }
      : {}),
    ...((override.shading ?? base.shading) ? { shading: override.shading ?? base.shading } : {}),
    ...((override.runProperties ?? base.runProperties)
      ? { runProperties: mergeRunPartial(base.runProperties ?? {}, override.runProperties ?? {}) }
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
