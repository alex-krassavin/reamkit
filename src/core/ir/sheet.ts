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

import type { Chart, DocumentInfo } from '@/core/document-model';
import type { DefinedName, ParsedWorksheet, XlsxStyles } from '@/core/spreadsheet-model';
import type { ResourceStore } from '@/core/ir/resources';

// §20.5 SpreadsheetDrawingML — a chart frame anchored over a sheet's grid,
// already sized from its two-cell anchor. chartPartPath keys into
// SheetDoc.chartData for the resolved chart.
export interface SheetChartRef {
  readonly chartPartPath: string;
  readonly widthPt: number;
  readonly heightPt: number;
}

export interface Sheet {
  readonly name: string;
  // The grid + per-sheet geometry exactly as parsed: cells, columns, rows,
  // merges, dimensions, pageSetup/printOptions, manual breaks, drawingRelId.
  readonly grid: ParsedWorksheet;
  // Chart frames on this sheet, anchor-ordered (resolved data in chartData).
  readonly charts?: ReadonlyArray<SheetChartRef>;
}

export interface SheetDoc {
  readonly kind: 'sheet';
  readonly sheets: ReadonlyArray<Sheet>;
  /** §18.8 workbook style table (cellXfs + fonts/fills/borders/numFmts). */
  readonly styles: XlsxStyles;
  /** §18.4 shared string table; a cell with type 's' indexes into it. */
  readonly sharedStrings: ReadonlyArray<string>;
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
