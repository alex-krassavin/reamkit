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

/**
 * §20.5 SpreadsheetDrawingML — a chart frame anchored over a sheet's grid,
 * already sized from its two-cell anchor. `chartPartPath` keys into
 * {@link SheetDoc.chartData} for the resolved chart.
 */
export interface SheetChartRef {
  readonly chartPartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
}

/**
 * §20.5.2.1 xdr:pic — a picture anchored over a sheet's grid, sized from its
 * anchor. The bytes live in {@link SheetDoc.resources}; `resourceId` keys into
 * it. The projection emits an ImageBlock per picture after the grid.
 */
export interface SheetImageRef {
  readonly resourceId: ResourceId;
  readonly widthPt: number;
  readonly heightPt: number;
}

/** One item (label + selection state) in a {@link SheetSlicer}. */
export interface SheetSlicerItem {
  /** The item's display label. */
  readonly label: string;
  /** Whether the item is currently selected (filtered in). */
  readonly selected: boolean;
}

/**
 * A slicer panel (xl/slicers + xl/slicerCaches, E-SHEET SV2) resolved against
 * its source so the projection can render it as a captioned button box after the
 * grid. Items + selection resolve for native-table slicers; an OLAP/pivot slicer
 * the cache cannot resolve inline degrades to a caption-only box.
 */
export interface SheetSlicer {
  readonly caption: string;
  readonly columnCount: number;
  readonly items: ReadonlyArray<SheetSlicerItem>;
  /**
   * Resolved fills + header text colour from the slicer style name + workbook
   * theme (the same accent heuristic as tables/pivots). Absent ⇒ a plain box.
   */
  readonly headerHex?: string;
  readonly selectedHex?: string;
  readonly headerTextHex?: string;
  readonly selectedTextHex?: string;
}

/**
 * §18.3.1.47 — a cell (or range) hyperlink resolved to an external URL (E-SHEET
 * W3). The projection stamps `run.href` on every cell the range covers.
 */
export interface SheetHyperlink {
  readonly ref: MergedRange;
  readonly url: string;
}

/**
 * A cell comment / note (E-SHEET W7): an A1 anchor, the resolved author display
 * name, the text body, and whether it came from the modern threaded comments
 * part (a conversation) or the legacy note part.
 */
export interface SheetComment {
  readonly ref: string;
  readonly author?: string;
  readonly text: string;
  readonly threaded: boolean;
}

/**
 * A form control resolved against its ctrlProp part (E-SHEET W8): the display
 * name, the `objectType` (CheckBox / Radio / Spin / …) and the bit of state the
 * listing shows — `checked` for check/option buttons, `value` for
 * spin/scroll/list controls. Render-only.
 */
export interface SheetFormControl {
  readonly name?: string;
  readonly objectType?: string;
  readonly checked?: boolean;
  readonly value?: number;
}

/**
 * An ActiveX control resolved against its activeX part (E-SHEET W10): the control
 * class (`type`, from the `<oleObject progId>`) plus the visible state persisted
 * in the property bag — `caption`, `value` (checked/text/number, as a string)
 * and OptionButton `groupName`. Render-only.
 */
export interface SheetActiveXControl {
  readonly type: string;
  readonly caption?: string;
  readonly value?: string;
  readonly groupName?: string;
}

/** One worksheet: the parsed grid plus its resolved, anchor-ordered overlays. */
export interface Sheet {
  /** The sheet (tab) name. */
  readonly name: string;
  /**
   * The grid + per-sheet geometry exactly as parsed: cells, columns, rows,
   * merges, dimensions, pageSetup/printOptions, manual breaks, drawingRelId.
   */
  readonly grid: ParsedWorksheet;
  /** Chart frames on this sheet, anchor-ordered (resolved data in {@link SheetDoc.chartData}). */
  readonly charts?: ReadonlyArray<SheetChartRef>;
  /** Picture frames on this sheet, anchor-ordered (bytes in {@link SheetDoc.resources}). */
  readonly images?: ReadonlyArray<SheetImageRef>;
  /** Drawing shapes on this sheet (E-SHEET W2), fully resolved + anchor-ordered. */
  readonly shapes?: ReadonlyArray<ShapeBlock>;
  /**
   * Cell hyperlinks resolved to external URLs (E-SHEET W3); the projection sets
   * `run.href` on covered cells. In-workbook (location-only) links are not carried.
   */
  readonly hyperlinks?: ReadonlyArray<SheetHyperlink>;
  /** Slicer panels on this sheet (E-SHEET SV2), rendered after the grid + charts. */
  readonly slicers?: ReadonlyArray<SheetSlicer>;
  /**
   * Cell comments / notes (E-SHEET W7), legacy and threaded. The projection
   * lists them in a "Comments" section after the grid (mirroring Excel's "print
   * comments at end of sheet"); render-only.
   */
  readonly comments?: ReadonlyArray<SheetComment>;
  /**
   * Form controls on this sheet (E-SHEET W8), resolved to type + state. Listed
   * in a "Form controls" section after the grid; render-only.
   */
  readonly formControls?: ReadonlyArray<SheetFormControl>;
  /**
   * ActiveX controls on this sheet (E-SHEET W10), resolved to type + visible
   * state. Listed in an "ActiveX controls" section after the grid; render-only.
   */
  readonly activeXControls?: ReadonlyArray<SheetActiveXControl>;
}

/**
 * The SpreadsheetML IR tree (E-SHEET): a workbook of grid {@link Sheet}s plus
 * its workbook-scoped companions — the style table, shared-string table, defined
 * names, date system, parsed charts, binary resources and metadata. Cells keep
 * their RAW stored value + a style index; resolution (shared strings, number
 * formats, the style cascade, the print model) happens at the projection
 * boundary, never in the tree.
 */
export interface SheetDoc {
  /** Discriminant for {@link SourceDoc} (projected to a FlowDoc for rendering). */
  readonly kind: 'sheet';
  /** The workbook's worksheets, in tab order. */
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
