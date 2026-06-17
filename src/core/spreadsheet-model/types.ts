// SpreadsheetML model (ECMA-376 Part 1 §18) — the format-neutral data shapes a
// worksheet and the workbook style table reduce to, the spreadsheet sibling of
// document-model. The xlsx parsers (src/excel) PRODUCE these; SheetDoc
// (core/ir/sheet) composes them; the print model projects them to a FlowDoc.
// Cells keep their RAW stored value + a style index — resolution (shared
// strings, number formats, the cascade) stays in the projection.

// ECMA-376 §18.3.1.4 — <c t="..."> cell value type.
export type CellType = 'n' | 's' | 'str' | 'b' | 'd' | 'e' | 'inlineStr';

export interface WorksheetCell {
  readonly column: number;
  readonly row: number;
  readonly type: CellType;
  // Raw stored value; renderer/converter looks up shared strings + formats.
  readonly rawValue: string;
  // For inlineStr the stored text lives in <is><t> rather than <v>.
  readonly inlineText?: string;
  // Index into the workbook's cellXfs (xl/styles.xml). 0 means default style.
  readonly styleIndex?: number;
}

export interface ColumnWidth {
  readonly min: number; // 1-indexed in OOXML, kept as-is here
  readonly max: number;
  readonly widthChars: number;
}

export interface MergedRange {
  readonly startColumn: number;
  readonly startRow: number;
  readonly endColumn: number;
  readonly endRow: number;
}

// ECMA-376 Part 1 §18.3.1.73 — <row ht="...">. The ht attribute is measured
// in points (not twips); customHeight="1" means the user pinned the height
// explicitly. Without customHeight, the height is content-driven.
export interface RowHeight {
  readonly row: number; // 0-indexed
  readonly heightPt: number;
  readonly customHeight: boolean;
}

// ECMA-376 Part 1 §18.3.1.62 — <pageMargins>. All attributes are in inches.
export interface XlsxPageMargins {
  readonly leftInches: number;
  readonly rightInches: number;
  readonly topInches: number;
  readonly bottomInches: number;
  readonly headerInches?: number;
  readonly footerInches?: number;
}

// ECMA-376 Part 1 §18.3.1.63 — <pageSetup>. paperSize is a numeric id from
// the printer paper size enumeration (1=Letter, 9=A4, ...). orientation
// values: 'default' (effectively portrait), 'portrait', 'landscape'.
//   scale       — print scaling percentage (10..400); default 100.
//   fitToWidth  — number of pages wide to fit to (0 ⇒ use scale); default 1.
//   fitToHeight — number of pages tall to fit to; default 1.
// fitToWidth/fitToHeight only take effect when <pageSetUpPr fitToPage="1">.
export interface XlsxPageSetup {
  readonly paperSize?: number;
  readonly orientation?: 'portrait' | 'landscape' | 'default';
  readonly scale?: number;
  readonly fitToWidth?: number;
  readonly fitToHeight?: number;
}

// ECMA-376 Part 1 §18.3.1.70 — <printOptions>. Controls what is rendered when
// the sheet is printed. gridLines defaults to false (Excel/Calc do NOT print
// cell gridlines unless explicitly enabled), so a faithful print model draws
// only the borders that come from cell styles.
export interface XlsxPrintOptions {
  readonly gridLines?: boolean;
  readonly horizontalCentered?: boolean;
  readonly verticalCentered?: boolean;
}

export interface ParsedWorksheet {
  readonly cells: ReadonlyArray<WorksheetCell>;
  readonly maxRow: number;
  readonly maxColumn: number;
  readonly columns: ReadonlyArray<ColumnWidth>;
  readonly merges: ReadonlyArray<MergedRange>;
  readonly rowHeights: ReadonlyArray<RowHeight>;
  readonly pageMargins?: XlsxPageMargins;
  readonly pageSetup?: XlsxPageSetup;
  // ECMA-376 §18.3.1.65 — <sheetPr><pageSetUpPr fitToPage="1"/>. When set, the
  // pageSetup fitToWidth/fitToHeight (not scale) drive print scaling.
  readonly fitToPage?: boolean;
  readonly printOptions?: XlsxPrintOptions;
  // ECMA-376 §18.3.1.74/§18.3.1.14 — manual <rowBreaks>/<colBreaks>. Each
  // stored value is the <brk id="..."> following the break (kept verbatim).
  readonly rowBreaks?: ReadonlyArray<number>;
  readonly colBreaks?: ReadonlyArray<number>;
  // §18.3.1.36 <drawing r:id> — the sheet's drawing part (charts/shapes).
  readonly drawingRelId?: string;
  // §18.3.1.18 <conditionalFormatting> — value-driven cell formats (E-SHEET SC1).
  readonly conditionalFormats?: ReadonlyArray<ConditionalFormat>;
  // §18.3.1.32 <dataValidations> — per-range input constraints (E-SHEET SV1). A
  // `list` validation paints an in-cell dropdown affordance; all types round-trip.
  readonly dataValidations?: ReadonlyArray<DataValidation>;
  // §18.3.1.47 <hyperlinks> — raw cell hyperlinks (E-SHEET W3). The reader
  // resolves each relId to an external URL; render-only (not written back).
  readonly hyperlinks?: ReadonlyArray<HyperlinkRef>;
  // §18.3.1.46 <headerFooter> — sheet header/footer format strings (E-SHEET W4).
  // The projection expands the &-codes into header/footer bands; render-only.
  readonly headerFooter?: HeaderFooter;
  // x14 extension <sparklineGroups> in extLst — per-cell mini charts (E-SHEET SC2).
  readonly sparklines?: ReadonlyArray<ParsedSparkline>;
  // §18.3.1.95 <tableParts> — relationship ids of the sheet's table parts. The
  // pure-XML parser only lists the ids; the reader resolves them (E-SHEET SC3).
  readonly tablePartRelIds?: ReadonlyArray<string>;
  // Resolved table parts (banded styles, header rows) — filled by the reader.
  readonly tables?: ReadonlyArray<ExcelTable>;
  // Resolved pivot tables (E-PIVOT) — discovered via the sheet's pivotTable
  // relationships (no element in the sheet XML), resolved by the reader.
  readonly pivotTables?: ReadonlyArray<PivotTable>;
  // ECMA-376 §18.3.1.66 <sheetView><pane state="frozen"> — frozen rows/columns.
  // A VIEW setting with NO effect on print/PDF (printed repeats are Print_Titles);
  // carried for round-trip fidelity and HTML sticky panes (E-SHEET SE2/SE3).
  readonly pane?: SheetPane;
}

// A frozen pane: the count of leading rows / columns frozen in the worksheet
// view. Derived from <pane ySplit="rows" xSplit="cols" state="frozen"> — a plain
// "split" pane (a resizable divider, no freeze) is not captured.
export interface SheetPane {
  readonly frozenRows: number;
  readonly frozenCols: number;
}

// ECMA-376 §18.5.1.2 <table> — a structured table over a cell range with a
// banded style. The raw parse carries the range, header rows and style flags;
// the reader resolves the named style to header / band fill colours against the
// workbook theme (E-SHEET SC3).
export interface ExcelTable {
  readonly ref: MergedRange;
  readonly name?: string;
  readonly styleName?: string;
  readonly headerRowCount: number;
  readonly showRowStripes: boolean;
  readonly showColumnStripes: boolean;
  readonly showFirstColumn: boolean;
  readonly showLastColumn: boolean;
  readonly autoFilter: boolean;
  // Resolved fills + header text colour (6-hex) — the reader derives these from
  // the named style + workbook theme (E-SHEET SC3).
  readonly headerHex?: string;
  readonly bandHex?: string;
  readonly headerTextHex?: string;
}

// ECMA-376 §18.10.1.73 <pivotTableDefinition> — a pivot table. Its OUTPUT cells
// are already cached in the worksheet, so they render as a normal grid; this
// carries the pivot's location and named style so the reader can band the region
// in the pivot's own palette (E-PIVOT). The structural model (row/column fields,
// subtotals, collapse state) is later work — only location + style here.
export interface PivotTable {
  readonly ref: MergedRange; // <location ref>
  readonly name?: string;
  readonly styleName?: string; // <pivotTableStyleInfo name>, e.g. PivotStyleDark1
  readonly firstHeaderRow: number; // header rows from the range top (default 1)
  readonly firstDataRow: number; // first data row, offset from the range top
  readonly firstDataCol: number; // first data column, offset from the range left
  readonly showRowStripes: boolean;
  readonly showColStripes: boolean;
  // §18.10.1.74 <rowItems>/<colItems> item type per data row / column (in
  // order): 'grand' (grand total), a subtotal-function name (subtotal), or
  // undefined (a plain data line). The i-th entry maps to data row/column
  // `firstDataRow/Col + i` (E-PIVOT PV3/PV4).
  readonly rowItemTypes?: ReadonlyArray<string | undefined>;
  readonly colItemTypes?: ReadonlyArray<string | undefined>;
  // Resolved fills + header text colour (6-hex) — derived from the named pivot
  // style + workbook theme by the reader (E-PIVOT PV2).
  readonly headerHex?: string;
  readonly bandHex?: string;
  readonly headerTextHex?: string;
}

// ECMA-376 Part 4 (x14 extension) — a sparkline: a mini chart drawn inside a
// single host cell (`sqref`) from a data range (`dataRange`, an A1 area possibly
// sheet-qualified). The group's type maps to line / column / win-loss.
export type SparklineKind = 'line' | 'column' | 'winLoss';

export interface ParsedSparkline {
  readonly kind: SparklineKind;
  readonly dataRange: string;
  readonly sqref: string;
  readonly colorHex?: string;
}

// ECMA-376 Part 1 §18.8 — the workbook style table (xl/styles.xml). Cells
// reference a cellXf by index; fonts/fills/borders/numFmts are looked up from
// it when the cell's properties are resolved.

export interface XlsxFont {
  readonly sizePt?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly colorHex?: string;
  readonly name?: string;
}

export interface XlsxFill {
  readonly patternType?: string;
  readonly fgColorHex?: string;
  readonly bgColorHex?: string;
}

export type XlsxBorderStyleName =
  | 'none'
  | 'thin'
  | 'medium'
  | 'thick'
  | 'hair'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface XlsxBorderEdge {
  readonly style?: XlsxBorderStyleName;
  readonly colorHex?: string;
}

export interface XlsxBorder {
  readonly top?: XlsxBorderEdge;
  readonly right?: XlsxBorderEdge;
  readonly bottom?: XlsxBorderEdge;
  readonly left?: XlsxBorderEdge;
  // §18.8.4 <diagonal> — the diagonal stroke; @diagonalUp (bottom-left → top-right)
  // and/or @diagonalDown (top-left → bottom-right) select which corners it spans (W6).
  readonly diagonal?: XlsxBorderEdge;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;
}

export type XlsxHorizontalAlign =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

export type XlsxVerticalAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface XlsxCellAlignment {
  readonly horizontal?: XlsxHorizontalAlign;
  readonly vertical?: XlsxVerticalAlign;
  readonly wrapText?: boolean;
  // §18.8.1 — left indent in "indent levels" (each ≈ 3 character widths) (W6).
  readonly indent?: number;
  // §18.8.1 textRotation — degrees counter-clockwise (0–90), 91–180 clockwise as
  // (value − 90), and 255 = stacked vertical text (W6).
  readonly textRotation?: number;
  // §18.8.1 shrinkToFit — scale the text down so it fits the cell on one line (W6).
  readonly shrinkToFit?: boolean;
}

export interface XlsxCellXf {
  readonly numFmtId: number;
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly applyNumberFormat?: boolean;
  readonly applyFont?: boolean;
  readonly applyFill?: boolean;
  readonly applyBorder?: boolean;
  readonly applyAlignment?: boolean;
  readonly alignment?: XlsxCellAlignment;
}

export interface XlsxStyles {
  readonly numFmts: ReadonlyMap<number, string>;
  readonly fonts: ReadonlyArray<XlsxFont>;
  readonly fills: ReadonlyArray<XlsxFill>;
  readonly borders: ReadonlyArray<XlsxBorder>;
  readonly cellXfs: ReadonlyArray<XlsxCellXf>;
  // §18.8.10 <dxfs> — differential formats referenced by conditional-format
  // rules (E-SHEET SC1); only the properties a dxf sets override the base.
  readonly dxfs?: ReadonlyArray<Dxf>;
}

// §18.8.14 <dxf> — a differential (override) format a cfRule applies on match.
export interface Dxf {
  readonly font?: XlsxFont;
  readonly fill?: XlsxFill;
}

// ECMA-376 Part 1 §18.2.5 — <definedName>. A workbook-scoped (or sheet-scoped,
// via localSheetId) named range. Print areas and print titles ride these under
// the reserved names _xlnm.Print_Area / _xlnm.Print_Titles.
export interface DefinedName {
  readonly name: string;
  readonly localSheetId?: number;
  readonly value: string;
}

// ECMA-376 Part 1 §18.3.1 — conditional formatting (E-SHEET SC1). A
// <conditionalFormatting sqref="…"> owns one or more <cfRule>s evaluated per
// cell in its range(s); the highest-priority matching rule wins.

// §18.18.15 ST_CfvoType — comparison operator for a `cellIs` rule.
export type CfOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

// §18.3.1.10 <cfRule type="cellIs"> — compares each cell to one constant (or
// two, for between/notBetween); a match applies the differential format dxfId.
export interface CfRuleCellIs {
  readonly type: 'cellIs';
  readonly priority: number;
  readonly operator: CfOperator;
  readonly formulas: ReadonlyArray<string>;
  readonly dxfId: number;
}

// §18.3.1.11 ST_CfvoType — how a <cfvo> stop's threshold is derived. `min`/`max`
// (and the dataBar `autoMin`/`autoMax`) take the range's extent; `num`/`formula`
// a literal; `percent`/`percentile` position within the value distribution.
export type CfvoType =
  | 'num'
  | 'percent'
  | 'max'
  | 'min'
  | 'percentile'
  | 'formula'
  | 'autoMin'
  | 'autoMax';

// §18.3.1.11 <cfvo> — one stop of a colorScale (or dataBar/iconSet). `val`
// carries the number/percent/formula text; absent for `min`/`max`.
export interface Cfvo {
  readonly type: CfvoType;
  readonly val?: string;
}

// §18.3.1.16 <cfRule type="colorScale"> — a 2- or 3-stop gradient. Each cfvo
// pairs with a colour; a cell's value, positioned between the bracketing stops,
// interpolates (in RGB) to a solid fill. Unlike cellIs it needs the range's
// value extent, so it resolves against every covered cell, not one constant.
export interface CfRuleColorScale {
  readonly type: 'colorScale';
  readonly priority: number;
  readonly cfvos: ReadonlyArray<Cfvo>;
  readonly colorsHex: ReadonlyArray<string>;
}

// §18.3.1.28 <cfRule type="dataBar"> — an in-cell bar whose length encodes the
// cell value within the range's extent. Two cfvo stops (lower/upper) bound the
// scale; `colorHex` fills the bar. minLength/maxLength clamp the bar as a percent
// (0..100) of the cell width (ECMA defaults 10/90; we default 0/100 — modern
// solid bars span the full cell).
export interface CfRuleDataBar {
  readonly type: 'dataBar';
  readonly priority: number;
  readonly cfvos: ReadonlyArray<Cfvo>;
  readonly colorHex: string;
  readonly minLength?: number;
  readonly maxLength?: number;
}

// §18.3.1.49 <cfRule type="iconSet"> — picks one glyph per cell from a named
// icon family (3/4/5 icons) by the value's bucket among the cfvo thresholds.
// `reverse` flips the icon order (highest value → first icon).
export interface CfRuleIconSet {
  readonly type: 'iconSet';
  readonly priority: number;
  readonly iconSet: string;
  readonly cfvos: ReadonlyArray<Cfvo>;
  readonly reverse?: boolean;
}

// §18.3.1.10 <cfRule type="top10"> — the top (or `bottom`) N values of the range
// take the differential format. `rank` is N; with `percent` it is a percentage of
// the range's cell count. Resolves against the range's value extent, like a scale.
export interface CfRuleTop10 {
  readonly type: 'top10';
  readonly priority: number;
  readonly rank: number;
  readonly percent: boolean;
  readonly bottom: boolean;
  readonly dxfId: number;
}

// §18.3.1.10 <cfRule type="aboveAverage"> — cells above (default) or below the
// range mean take the format. `equalAverage` makes the comparison inclusive;
// `stdDev`, when set, shifts the threshold by N population standard deviations.
export interface CfRuleAboveAverage {
  readonly type: 'aboveAverage';
  readonly priority: number;
  readonly aboveAverage: boolean;
  readonly equalAverage: boolean;
  readonly stdDev?: number;
  readonly dxfId: number;
}

// §18.3.1.10 <cfRule type="duplicateValues" | "uniqueValues"> — cells whose value
// repeats within the range (duplicate) or occurs exactly once (unique) take the
// format. Compares numbers by value and strings case-insensitively, like Excel.
export interface CfRuleDupUnique {
  readonly type: 'duplicateValues' | 'uniqueValues';
  readonly priority: number;
  readonly dxfId: number;
}

// §18.3.1.10 <cfRule type="containsText" | "notContainsText" | "beginsWith" |
// "endsWith"> — a case-insensitive substring test against the cell's text.
// `text` is the needle; `formula` carries Excel's generated SEARCH/LEFT/RIGHT
// expression verbatim for faithful write-back (it is matched directly, not run).
export interface CfRuleText {
  readonly type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';
  readonly priority: number;
  readonly text: string;
  readonly dxfId: number;
  readonly formula?: string;
}

export type CfRule =
  | CfRuleCellIs
  | CfRuleColorScale
  | CfRuleDataBar
  | CfRuleIconSet
  | CfRuleTop10
  | CfRuleAboveAverage
  | CfRuleDupUnique
  | CfRuleText;

// §18.3.1.18 <conditionalFormatting sqref="A1:A10 C1:C5"> — rules over ranges.
export interface ConditionalFormat {
  readonly ranges: ReadonlyArray<MergedRange>;
  readonly rules: ReadonlyArray<CfRule>;
}

// §18.3.1.47 <hyperlink ref r:id location display tooltip> — a cell (or range)
// hyperlink (E-SHEET W3). `relId` resolves to an external URL through the
// worksheet relationships; `location` is an in-workbook target (Sheet!cell). The
// raw form rides on the grid; the reader resolves it to a SheetHyperlink.
export interface HyperlinkRef {
  readonly ref: string;
  readonly relId?: string;
  readonly location?: string;
  readonly display?: string;
  readonly tooltip?: string;
}

// §18.3.1.46 <headerFooter> — the sheet's print header/footer format strings
// (E-SHEET W4). Each carries Excel's &-code mini-language (&L/&C/&R regions,
// &P/&N/&D/&F/&A field codes, &B/&I formatting). v1 reads the odd (= default)
// header and footer; even/first variants are a later refinement.
export interface HeaderFooter {
  readonly oddHeader?: string;
  readonly oddFooter?: string;
}

// §18.4.4 <r> — one formatting run inside a rich-text shared string (E-SHEET W6).
// A `<si>` with multiple `<r><rPr>…</rPr><t>…</t></r>` runs carries per-run
// formatting from <rPr> (its own font properties, not a cellXf index). The
// projection emits these as separate document-model runs so a single cell can mix
// bold / colour / size. Render-only: the writer flattens back to plain text, so
// the round-trip stays byte-stable (rich formatting in a shared string is dropped
// on `convert('xlsx')`, a documented loss like slicers/pivots).
export interface SheetRichRun {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly colorHex?: string;
  readonly sizePt?: number;
  // §18.4.2 <vertAlign> — superscript / subscript within the cell text.
  readonly vertAlign?: 'superscript' | 'subscript';
}

// §18.18.18 ST_DataValidationType — the constraint a <dataValidation> enforces.
// Only `list` has a visual signature (the in-cell dropdown); the rest carry
// through for round-trip fidelity and for surfacing the input/error prompts.
export type DataValidationType =
  | 'none'
  | 'whole'
  | 'decimal'
  | 'list'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom';

// §18.3.1.33 <dataValidation> — an input constraint over one or more ranges
// (sqref) (E-SHEET SV1). The visually-meaningful part is `type` (a `list` cell
// shows a dropdown); the formulas + prompts ride through so the validation
// survives a read→write round-trip. `showDropDown` keeps ECMA's INVERTED sense:
// the attribute is "1" to HIDE the in-cell dropdown, so a list validation shows
// one when the flag is absent/false.
export interface DataValidation {
  readonly type: DataValidationType;
  readonly ranges: ReadonlyArray<MergedRange>;
  readonly operator?: string;
  readonly allowBlank?: boolean;
  readonly showDropDown?: boolean;
  readonly showInputMessage?: boolean;
  readonly showErrorMessage?: boolean;
  readonly errorStyle?: string;
  readonly formula1?: string;
  readonly formula2?: string;
  readonly promptTitle?: string;
  readonly prompt?: string;
  readonly errorTitle?: string;
  readonly error?: string;
}
