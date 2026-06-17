// SheetDoc — the SpreadsheetML IR tree (E-SHEET), the spreadsheet sibling of
// FlowDoc. A workbook of grid sheets plus its workbook-scoped companions: the
// style table, the shared-string table, defined names, the date system, parsed
// charts, binary resources and metadata.
//
// Cells keep their RAW stored value + a style index — resolution (shared
// strings, number formats, the style cascade, the print model) happens at the
// projection boundary (SheetDoc → FlowDoc today, a dedicated grid layout
// later), never in the tree. Like FlowDoc, caller-supplied conversion options
// (fonts, PDF/A profile, …) are NOT part of it — they parameterize transforms.

import type { Chart, DocumentInfo, ShapeBlock } from '@/core/document-model';
import type {
  DefinedName,
  MergedRange,
  ParsedWorksheet,
  SheetRichRun,
  XlsxStyles,
} from '@/core/spreadsheet-model';
import type { ResourceId, ResourceStore } from '@/core/ir/resources';

// §20.5 SpreadsheetDrawingML — a chart frame anchored over a sheet's grid,
// already sized from its two-cell anchor. chartPartPath keys into
// SheetDoc.chartData for the resolved chart.
export interface SheetChartRef {
  readonly chartPartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
}

// §20.5.2.1 xdr:pic — a picture anchored over a sheet's grid, sized from its
// anchor. The bytes live in SheetDoc.resources; `resourceId` keys into it. The
// projection emits an ImageBlock per picture after the grid (like chart frames).
export interface SheetImageRef {
  readonly resourceId: ResourceId;
  readonly widthPt: number;
  readonly heightPt: number;
}

// A slicer panel (xl/slicers + xl/slicerCaches, E-SHEET SV2) resolved against
// its source so the projection can render it as a captioned button box after the
// grid (mirroring how chart frames render after it). Items + selection resolve
// for native-table slicers (the cache's tableSlicerCache → the table column's
// distinct values, with the autofilter giving selection); an OLAP/pivot slicer
// the cache cannot resolve inline degrades to a caption-only box (no items).
export interface SheetSlicerItem {
  readonly label: string;
  readonly selected: boolean;
}

export interface SheetSlicer {
  readonly caption: string;
  readonly columnCount: number;
  readonly items: ReadonlyArray<SheetSlicerItem>;
  // Resolved fills + header text colour from the slicer style name + workbook
  // theme (the same accent heuristic as tables/pivots). Absent ⇒ a plain box.
  readonly headerHex?: string;
  readonly selectedHex?: string;
  readonly headerTextHex?: string;
  readonly selectedTextHex?: string;
}

// §18.3.1.47 — a cell (or range) hyperlink resolved to an external URL (E-SHEET
// W3). The projection stamps run.href on every cell the range covers.
export interface SheetHyperlink {
  readonly ref: MergedRange;
  readonly url: string;
}

export interface Sheet {
  readonly name: string;
  // The grid + per-sheet geometry exactly as parsed: cells, columns, rows,
  // merges, dimensions, pageSetup/printOptions, manual breaks, drawingRelId.
  readonly grid: ParsedWorksheet;
  // Chart frames on this sheet, anchor-ordered (resolved data in chartData).
  readonly charts?: ReadonlyArray<SheetChartRef>;
  // Picture frames on this sheet, anchor-ordered (bytes in SheetDoc.resources).
  readonly images?: ReadonlyArray<SheetImageRef>;
  // Drawing shapes on this sheet (E-SHEET W2), fully resolved + anchor-ordered.
  readonly shapes?: ReadonlyArray<ShapeBlock>;
  // Cell hyperlinks resolved to external URLs (E-SHEET W3); the projection sets
  // run.href on covered cells. In-workbook (location-only) links are not carried.
  readonly hyperlinks?: ReadonlyArray<SheetHyperlink>;
  // Slicer panels on this sheet (E-SHEET SV2), rendered after the grid + charts.
  readonly slicers?: ReadonlyArray<SheetSlicer>;
}

export interface SheetDoc {
  readonly kind: 'sheet';
  readonly sheets: ReadonlyArray<Sheet>;
  /** §18.8 workbook style table (cellXfs + fonts/fills/borders/numFmts). */
  readonly styles: XlsxStyles;
  /** §18.4 shared string table; a cell with type 's' indexes into it. */
  readonly sharedStrings: ReadonlyArray<string>;
  /**
   * §18.4.4 per-index rich-text runs (E-SHEET W6), parallel to sharedStrings —
   * `[i]` is defined only when shared string i carries per-run formatting.
   * Render-only: the writer flattens to plain text, so it is absent on round-trip.
   */
  readonly sharedStringRuns?: ReadonlyArray<ReadonlyArray<SheetRichRun> | undefined>;
  /** §18.2.5 workbook defined names (print areas/titles, named ranges). */
  readonly definedNames: ReadonlyArray<DefinedName>;
  /** §18.2.28 1904 date system (serial-to-date epoch). */
  readonly date1904: boolean;
  /** Parsed charts keyed by part path (SheetChartRef.chartPartPath). */
  readonly chartData?: ReadonlyMap<string, Chart>;
  /** Content-addressed binary resources (sheet images). */
  readonly resources: ResourceStore;
  /** Document metadata from docProps/core.xml. */
  readonly info?: DocumentInfo;
}
