// Build a minimal valid xlsx package in memory for tests, optionally with a
// custom styles.xml and merged cell ranges.

import { zipSync } from 'fflate';

import { formatCellRef } from '@/excel';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const CONTENT_TYPES_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;

const CONTENT_TYPES_FIXED_OVERRIDES =
  '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';

const STYLES_TYPE_OVERRIDE =
  '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

export type XlsxValue = string | number | boolean | null;

export interface XlsxCellSpec {
  readonly value: XlsxValue;
  readonly styleIndex?: number;
}

export interface XlsxRowHeightSpec {
  readonly row: number; // 0-indexed
  readonly heightPt: number;
  readonly customHeight?: boolean;
}

export interface XlsxPageMarginsSpec {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
  readonly header?: number;
  readonly footer?: number;
}

export interface XlsxPageSetupSpec {
  readonly paperSize?: number;
  readonly orientation?: 'portrait' | 'landscape';
  readonly scale?: number;
  readonly fitToWidth?: number;
  readonly fitToHeight?: number;
}

export interface XlsxPrintOptionsSpec {
  readonly gridLines?: boolean;
  readonly horizontalCentered?: boolean;
  readonly verticalCentered?: boolean;
}

export interface XlsxDefinedNameSpec {
  readonly name: string;
  readonly localSheetId?: number;
  readonly value: string;
}

export interface XlsxSheetSpec {
  readonly name: string;
  readonly rows: ReadonlyArray<ReadonlyArray<XlsxValue | XlsxCellSpec>>;
  readonly mergeRefs?: ReadonlyArray<string>;
  readonly columns?: ReadonlyArray<{
    readonly min: number;
    readonly max: number;
    readonly widthChars: number;
  }>;
  readonly rowHeights?: ReadonlyArray<XlsxRowHeightSpec>;
  readonly pageMargins?: XlsxPageMarginsSpec;
  readonly pageSetup?: XlsxPageSetupSpec;
  readonly printOptions?: XlsxPrintOptionsSpec;
  readonly fitToPage?: boolean;
  readonly rowBreaks?: ReadonlyArray<number>;
  readonly colBreaks?: ReadonlyArray<number>;
  /** <sheetView><pane state="frozen"> — frozen leading rows / columns. */
  readonly freeze?: { readonly rows?: number; readonly cols?: number };
  /** Raw <conditionalFormatting> markup injected into the worksheet. */
  readonly conditionalFormattingXml?: string;
  /** Raw <dataValidations> markup injected into the worksheet (E-SHEET SV1). */
  readonly dataValidationsXml?: string;
  /** Raw <extLst> markup injected at the end of the worksheet (x14 sparklines). */
  readonly extLstXml?: string;
}

export interface XlsxBuilderOptions {
  readonly rows?: ReadonlyArray<ReadonlyArray<XlsxValue | XlsxCellSpec>>;
  readonly sheets?: ReadonlyArray<XlsxSheetSpec>;
  readonly stylesXml?: string;
  readonly mergeRefs?: ReadonlyArray<string>;
  readonly columns?: ReadonlyArray<{
    readonly min: number;
    readonly max: number;
    readonly widthChars: number;
  }>;
  readonly rowHeights?: ReadonlyArray<XlsxRowHeightSpec>;
  readonly pageMargins?: XlsxPageMarginsSpec;
  readonly pageSetup?: XlsxPageSetupSpec;
  readonly printOptions?: XlsxPrintOptionsSpec;
  readonly fitToPage?: boolean;
  readonly rowBreaks?: ReadonlyArray<number>;
  readonly colBreaks?: ReadonlyArray<number>;
  readonly freeze?: { readonly rows?: number; readonly cols?: number };
  readonly conditionalFormattingXml?: string;
  readonly dataValidationsXml?: string;
  readonly extLstXml?: string;
  readonly date1904?: boolean;
  readonly definedNames?: ReadonlyArray<XlsxDefinedNameSpec>;
  /** Attach a chart to the FIRST sheet via a drawing part (twoCellAnchor). */
  readonly sheetChart?: {
    readonly chartXml: string; // full c:chartSpace markup
    readonly colorsXml?: string; // full cs:colorStyle markup (colors1.xml)
    /** Anchor cells; defaults to B2..H17 (≈ 6×4 inches on default tracks). */
    readonly anchor?: { from: [number, number]; to: [number, number] };
  };
  /** Attach a picture to the FIRST sheet via a drawing part (xdr:pic, W1). */
  readonly sheetImage?: {
    readonly pngBytes: Uint8Array;
    readonly anchor?: { from: [number, number]; to: [number, number] };
  };
  /** Attach a shape to the FIRST sheet via a drawing part (xdr:sp, W2). */
  readonly sheetShape?: {
    readonly text?: string;
    readonly fillHex?: string;
    readonly preset?: string;
    readonly anchor?: { from: [number, number]; to: [number, number] };
  };
  /** Attach cell hyperlinks to the FIRST sheet (E-SHEET W3). A `url` is emitted as
   *  an external `r:id` relationship; a `location` is an in-workbook target. */
  readonly hyperlinks?: ReadonlyArray<{
    readonly ref: string;
    readonly url?: string;
    readonly location?: string;
    readonly tooltip?: string;
  }>;
  /** Attach Excel table parts to the FIRST sheet (E-SHEET SC3). */
  readonly tables?: ReadonlyArray<{
    readonly ref: string;
    readonly name?: string;
    readonly styleName?: string;
    readonly showRowStripes?: boolean;
    readonly headerRowCount?: number;
    /** <autoFilter><filterColumn colId><filters><filter val> — slicer selection. */
    readonly filters?: ReadonlyArray<{
      readonly colId: number;
      readonly values: ReadonlyArray<string>;
    }>;
  }>;
  /** Attach slicer + slicerCache parts (E-SHEET SV2). The slicer part rel is on
   *  the FIRST sheet; the cache rel is on the workbook. A `cache.tableId` binds
   *  the slicer to a `tables` entry (1-based by order) for native-table items. */
  readonly slicers?: ReadonlyArray<{
    readonly name: string;
    readonly caption: string;
    readonly cacheName: string;
    readonly columnCount?: number;
    readonly styleName?: string;
    readonly cache: {
      readonly sourceName?: string;
      readonly tableId?: number;
      readonly column?: number; // table column id (1-based)
    };
  }>;
  /** Attach pivot table parts to the FIRST sheet (E-PIVOT). Referenced via a
   *  sheet relationship only — no element in the worksheet XML. */
  readonly pivotTables?: ReadonlyArray<{
    readonly ref: string;
    readonly name?: string;
    readonly styleName?: string;
    readonly firstHeaderRow?: number;
    readonly firstDataRow?: number;
    readonly firstDataCol?: number;
    readonly showRowStripes?: boolean;
    readonly showColStripes?: boolean;
    /** <rowItems>/<colItems> @t per data row/column ('grand'/subtotal/undefined). */
    readonly rowItemTypes?: ReadonlyArray<string | undefined>;
    readonly colItemTypes?: ReadonlyArray<string | undefined>;
  }>;
}

function isCellSpec(v: XlsxValue | XlsxCellSpec): v is XlsxCellSpec {
  return typeof v === 'object' && v !== null && 'value' in v;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildTableXml(
  id: number,
  t: {
    ref: string;
    name?: string;
    styleName?: string;
    showRowStripes?: boolean;
    headerRowCount?: number;
    filters?: ReadonlyArray<{ colId: number; values: ReadonlyArray<string> }>;
  },
): string {
  const name = t.name ?? `Table${id}`;
  const style = t.styleName ?? 'TableStyleMedium2';
  const stripes = t.showRowStripes === false ? '0' : '1';
  const hrc = t.headerRowCount !== undefined ? ` headerRowCount="${t.headerRowCount}"` : '';
  const autoFilter =
    t.filters && t.filters.length > 0
      ? `<autoFilter ref="${t.ref}">${t.filters
          .map(
            (fc) =>
              `<filterColumn colId="${fc.colId}"><filters>${fc.values
                .map((v) => `<filter val="${escapeXml(v)}"/>`)
                .join('')}</filters></filterColumn>`,
          )
          .join('')}</autoFilter>`
      : `<autoFilter ref="${t.ref}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${name}" displayName="${name}" ref="${t.ref}"${hrc} totalsRowShown="0">
  ${autoFilter}
  <tableColumns count="1"><tableColumn id="1" name="Col"/></tableColumns>
  <tableStyleInfo name="${style}" showFirstColumn="0" showLastColumn="0" showRowStripes="${stripes}" showColumnStripes="0"/>
</table>`;
}

function buildPivotTableXml(
  id: number,
  p: {
    ref: string;
    name?: string;
    styleName?: string;
    firstHeaderRow?: number;
    firstDataRow?: number;
    firstDataCol?: number;
    showRowStripes?: boolean;
    showColStripes?: boolean;
    rowItemTypes?: ReadonlyArray<string | undefined>;
    colItemTypes?: ReadonlyArray<string | undefined>;
  },
): string {
  const name = p.name ?? `PivotTable${id}`;
  const style = p.styleName ?? 'PivotStyleLight16';
  const fhr = p.firstHeaderRow ?? 1;
  const fdr = p.firstDataRow ?? 2;
  const fdc = p.firstDataCol ?? 1;
  const rowStripes = p.showRowStripes ? '1' : '0';
  const colStripes = p.showColStripes ? '1' : '0';
  const items = (tag: string, types?: ReadonlyArray<string | undefined>): string =>
    types
      ? `<${tag} count="${types.length}">${types
          .map((t) => `<i${t ? ` t="${t}"` : ''}><x/></i>`)
          .join('')}</${tag}>`
      : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${name}" cacheId="0" dataCaption="Values" outline="1" outlineData="1">
  <location ref="${p.ref}" firstHeaderRow="${fhr}" firstDataRow="${fdr}" firstDataCol="${fdc}"/>
  ${items('rowItems', p.rowItemTypes)}${items('colItems', p.colItemTypes)}<pivotTableStyleInfo name="${style}" showRowHeaders="1" showColHeaders="1" showRowStripes="${rowStripes}" showColStripes="${colStripes}" showLastColumn="0"/>
</pivotTableDefinition>`;
}

export function buildXlsx(
  rowsOrOptions: ReadonlyArray<ReadonlyArray<XlsxValue>> | XlsxBuilderOptions,
): Uint8Array {
  const options: XlsxBuilderOptions = Array.isArray(rowsOrOptions)
    ? { rows: rowsOrOptions as ReadonlyArray<ReadonlyArray<XlsxValue>> }
    : rowsOrOptions;

  const sheets: Array<XlsxSheetSpec> =
    options.sheets && options.sheets.length > 0
      ? [...options.sheets]
      : [
          {
            name: 'Sheet1',
            rows: options.rows ?? [],
            ...(options.mergeRefs ? { mergeRefs: options.mergeRefs } : {}),
            ...(options.columns ? { columns: options.columns } : {}),
            ...(options.rowHeights ? { rowHeights: options.rowHeights } : {}),
            ...(options.pageMargins ? { pageMargins: options.pageMargins } : {}),
            ...(options.pageSetup ? { pageSetup: options.pageSetup } : {}),
            ...(options.printOptions ? { printOptions: options.printOptions } : {}),
            ...(options.fitToPage !== undefined ? { fitToPage: options.fitToPage } : {}),
            ...(options.rowBreaks ? { rowBreaks: options.rowBreaks } : {}),
            ...(options.colBreaks ? { colBreaks: options.colBreaks } : {}),
            ...(options.freeze ? { freeze: options.freeze } : {}),
            ...(options.conditionalFormattingXml
              ? { conditionalFormattingXml: options.conditionalFormattingXml }
              : {}),
            ...(options.dataValidationsXml
              ? { dataValidationsXml: options.dataValidationsXml }
              : {}),
            ...(options.extLstXml ? { extLstXml: options.extLstXml } : {}),
          },
        ];

  const sharedStringsList: Array<string> = [];
  const sharedStringIndex = new Map<string, number>();
  const internString = (s: string): number => {
    const existing = sharedStringIndex.get(s);
    if (existing !== undefined) return existing;
    const idx = sharedStringsList.length;
    sharedStringIndex.set(s, idx);
    sharedStringsList.push(s);
    return idx;
  };

  const sheetParts: Array<{ fileName: string; xml: string }> = [];
  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s]!;
    const rowHeightLookup = new Map<number, XlsxRowHeightSpec>();
    for (const h of sheet.rowHeights ?? []) rowHeightLookup.set(h.row, h);
    const sheetRows: Array<string> = [];
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r]!;
      const cells: Array<string> = [];
      for (let c = 0; c < row.length; c++) {
        const item = row[c];
        const spec: XlsxCellSpec | undefined =
          item === null || item === undefined
            ? undefined
            : isCellSpec(item)
              ? item
              : { value: item };
        if (!spec || spec.value === null) continue;
        const ref = formatCellRef({ row: r, column: c });
        const styleAttr = spec.styleIndex !== undefined ? ` s="${spec.styleIndex}"` : '';
        if (typeof spec.value === 'number') {
          cells.push(`<c r="${ref}"${styleAttr}><v>${spec.value}</v></c>`);
        } else if (typeof spec.value === 'boolean') {
          cells.push(`<c r="${ref}"${styleAttr} t="b"><v>${spec.value ? 1 : 0}</v></c>`);
        } else {
          const idx = internString(spec.value);
          cells.push(`<c r="${ref}"${styleAttr} t="s"><v>${idx}</v></c>`);
        }
      }
      const heightSpec = rowHeightLookup.get(r);
      const heightAttrs = heightSpec
        ? ` ht="${heightSpec.heightPt}"${heightSpec.customHeight !== false ? ' customHeight="1"' : ''}`
        : '';
      if (cells.length > 0 || heightSpec) {
        sheetRows.push(`  <row r="${r + 1}"${heightAttrs}>${cells.join('')}</row>`);
      }
    }

    const colsXml =
      sheet.columns && sheet.columns.length > 0
        ? '<cols>' +
          sheet.columns
            .map(
              (c) => `<col min="${c.min}" max="${c.max}" width="${c.widthChars}" customWidth="1"/>`,
            )
            .join('') +
          '</cols>'
        : '';
    const mergeXml =
      sheet.mergeRefs && sheet.mergeRefs.length > 0
        ? `<mergeCells count="${sheet.mergeRefs.length}">` +
          sheet.mergeRefs.map((r) => `<mergeCell ref="${escapeXml(r)}"/>`).join('') +
          '</mergeCells>'
        : '';
    const marginsXml = sheet.pageMargins
      ? `<pageMargins left="${sheet.pageMargins.left}" right="${sheet.pageMargins.right}"` +
        ` top="${sheet.pageMargins.top}" bottom="${sheet.pageMargins.bottom}"` +
        (sheet.pageMargins.header !== undefined ? ` header="${sheet.pageMargins.header}"` : '') +
        (sheet.pageMargins.footer !== undefined ? ` footer="${sheet.pageMargins.footer}"` : '') +
        '/>'
      : '';
    const setupXml = sheet.pageSetup
      ? `<pageSetup` +
        (sheet.pageSetup.paperSize !== undefined
          ? ` paperSize="${sheet.pageSetup.paperSize}"`
          : '') +
        (sheet.pageSetup.orientation !== undefined
          ? ` orientation="${sheet.pageSetup.orientation}"`
          : '') +
        (sheet.pageSetup.scale !== undefined ? ` scale="${sheet.pageSetup.scale}"` : '') +
        (sheet.pageSetup.fitToWidth !== undefined
          ? ` fitToWidth="${sheet.pageSetup.fitToWidth}"`
          : '') +
        (sheet.pageSetup.fitToHeight !== undefined
          ? ` fitToHeight="${sheet.pageSetup.fitToHeight}"`
          : '') +
        '/>'
      : '';
    const sheetPrXml = sheet.fitToPage ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : '';
    const printOptionsXml = sheet.printOptions
      ? `<printOptions` +
        (sheet.printOptions.gridLines !== undefined
          ? ` gridLines="${sheet.printOptions.gridLines ? 1 : 0}"`
          : '') +
        (sheet.printOptions.horizontalCentered !== undefined
          ? ` horizontalCentered="${sheet.printOptions.horizontalCentered ? 1 : 0}"`
          : '') +
        (sheet.printOptions.verticalCentered !== undefined
          ? ` verticalCentered="${sheet.printOptions.verticalCentered ? 1 : 0}"`
          : '') +
        '/>'
      : '';
    const rowBreaksXml =
      sheet.rowBreaks && sheet.rowBreaks.length > 0
        ? `<rowBreaks count="${sheet.rowBreaks.length}" manualBreakCount="${sheet.rowBreaks.length}">` +
          sheet.rowBreaks.map((id) => `<brk id="${id}" max="16383" man="1"/>`).join('') +
          '</rowBreaks>'
        : '';
    const colBreaksXml =
      sheet.colBreaks && sheet.colBreaks.length > 0
        ? `<colBreaks count="${sheet.colBreaks.length}" manualBreakCount="${sheet.colBreaks.length}">` +
          sheet.colBreaks.map((id) => `<brk id="${id}" max="1048575" man="1"/>`).join('') +
          '</colBreaks>'
        : '';
    const freezeCols = sheet.freeze?.cols ?? 0;
    const freezeRows = sheet.freeze?.rows ?? 0;
    const sheetViewsXml =
      freezeCols > 0 || freezeRows > 0
        ? '<sheetViews><sheetView workbookViewId="0"><pane' +
          (freezeCols > 0 ? ` xSplit="${freezeCols}"` : '') +
          (freezeRows > 0 ? ` ySplit="${freezeRows}"` : '') +
          ` state="frozen"/>` +
          '</sheetView></sheetViews>'
        : '';
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetPrXml}
  ${sheetViewsXml}
  ${colsXml}
  <sheetData>
${sheetRows.join('\n')}
  </sheetData>
  ${printOptionsXml}
  ${mergeXml}
  ${sheet.conditionalFormattingXml ?? ''}
  ${sheet.dataValidationsXml ?? ''}
  ${marginsXml}
  ${setupXml}
  ${rowBreaksXml}
  ${colBreaksXml}
  ${sheet.extLstXml ?? ''}
</worksheet>`;
    sheetParts.push({ fileName: `sheet${s + 1}.xml`, xml });
  }

  const workbookXml = buildWorkbookXml(
    sheets,
    options.date1904 ?? false,
    options.definedNames ?? [],
  );
  const workbookRelsXml = buildWorkbookRelsXml(sheets.length, Boolean(options.stylesXml));

  const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStringsList.length}" uniqueCount="${sharedStringsList.length}">
${sharedStringsList.map((str) => `  <si><t>${escapeXml(str)}</t></si>`).join('\n')}
</sst>`;

  const sheetOverrides = sheetParts
    .map(
      (s) =>
        `<Override PartName="/xl/worksheets/${s.fileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('');
  const contentTypes =
    CONTENT_TYPES_HEADER +
    sheetOverrides +
    CONTENT_TYPES_FIXED_OVERRIDES +
    (options.stylesXml ? STYLES_TYPE_OVERRIDE : '') +
    (options.sheetChart
      ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
        '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
        (options.sheetChart.colorsXml
          ? '<Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>'
          : '')
      : '') +
    (options.sheetImage
      ? '<Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      : '') +
    (options.sheetShape && !options.sheetImage
      ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      : '') +
    (options.tables
      ? options.tables
          .map(
            (_, i) =>
              `<Override PartName="/xl/tables/table${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`,
          )
          .join('')
      : '') +
    (options.pivotTables
      ? options.pivotTables
          .map(
            (_, i) =>
              `<Override PartName="/xl/pivotTables/pivotTable${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`,
          )
          .join('')
      : '') +
    (options.slicers
      ? options.slicers
          .map(
            (_, i) =>
              `<Override PartName="/xl/slicers/slicer${i + 1}.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>` +
              `<Override PartName="/xl/slicerCaches/slicerCache${i + 1}.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>`,
          )
          .join('')
      : '') +
    '\n</Types>';

  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(contentTypes),
    '_rels/.rels': encoder.encode(ROOT_RELS),
    'xl/workbook.xml': encoder.encode(workbookXml),
    'xl/_rels/workbook.xml.rels': encoder.encode(workbookRelsXml),
    'xl/sharedStrings.xml': encoder.encode(sharedStringsXml),
  };
  if (options.sheetChart && sheetParts.length > 0) {
    const first = sheetParts[0]!;
    first.xml = first.xml.replace('</worksheet>', '<drawing r:id="rId100"/></worksheet>');
    entries[`xl/worksheets/_rels/${first.fileName}.rels`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
    );
    const a = options.sheetChart.anchor ?? { from: [1, 1], to: [7, 16] };
    entries['xl/drawings/drawing1.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${a.from[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.from[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${a.to[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.to[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
    );
    const chartRels = options.sheetChart.colorsXml
      ? '  <Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>\n'
      : '';
    entries['xl/drawings/_rels/drawing1.xml.rels'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
    );
    entries['xl/charts/chart1.xml'] = encoder.encode(options.sheetChart.chartXml);
    if (options.sheetChart.colorsXml) {
      entries['xl/charts/_rels/chart1.xml.rels'] = encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${chartRels}</Relationships>`,
      );
      entries['xl/charts/colors1.xml'] = encoder.encode(options.sheetChart.colorsXml);
    }
  }
  if (options.sheetImage && sheetParts.length > 0) {
    const first = sheetParts[0]!;
    first.xml = first.xml.replace('</worksheet>', '<drawing r:id="rId100"/></worksheet>');
    entries[`xl/worksheets/_rels/${first.fileName}.rels`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
    );
    const a = options.sheetImage.anchor ?? { from: [1, 1], to: [4, 8] };
    entries['xl/drawings/drawing1.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${a.from[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.from[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${a.to[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.to[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
    );
    entries['xl/drawings/_rels/drawing1.xml.rels'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
    );
    entries['xl/media/image1.png'] = options.sheetImage.pngBytes;
  }
  if (options.sheetShape && sheetParts.length > 0) {
    const first = sheetParts[0]!;
    first.xml = first.xml.replace('</worksheet>', '<drawing r:id="rId100"/></worksheet>');
    entries[`xl/worksheets/_rels/${first.fileName}.rels`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
    );
    const a = options.sheetShape.anchor ?? { from: [1, 1], to: [4, 6] };
    const fill = options.sheetShape.fillHex ?? '4472C4';
    const preset = options.sheetShape.preset ?? 'roundRect';
    const text = options.sheetShape.text ?? 'Shape text';
    entries['xl/drawings/drawing1.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${a.from[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.from[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${a.to[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.to[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
        <a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln>
      </xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
    );
  }
  if (options.hyperlinks && options.hyperlinks.length > 0 && sheetParts.length > 0) {
    const first = sheetParts[0]!;
    const rels: Array<string> = [];
    const linkTags = options.hyperlinks.map((h, i) => {
      const tip = h.tooltip ? ` tooltip="${escapeXml(h.tooltip)}"` : '';
      if (h.url !== undefined) {
        const rid = `rIdHl${i + 1}`;
        rels.push(
          `  <Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(h.url)}" TargetMode="External"/>`,
        );
        return `<hyperlink ref="${escapeXml(h.ref)}" r:id="${rid}"${tip}/>`;
      }
      const loc = h.location ? ` location="${escapeXml(h.location)}"` : '';
      return `<hyperlink ref="${escapeXml(h.ref)}"${loc}${tip}/>`;
    });
    first.xml = first.xml.replace(
      '</worksheet>',
      `<hyperlinks>${linkTags.join('')}</hyperlinks></worksheet>`,
    );
    if (rels.length > 0) {
      const relsPath = `xl/worksheets/_rels/${first.fileName}.rels`;
      const existing = entries[relsPath] ? decoder.decode(entries[relsPath]) : undefined;
      entries[relsPath] = encoder.encode(
        existing
          ? existing.replace('</Relationships>', `${rels.join('\n')}\n</Relationships>`)
          : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join('\n')}
</Relationships>`,
      );
    }
  }
  if (options.tables && options.tables.length > 0 && sheetParts.length > 0) {
    const first = sheetParts[0]!;
    const parts = options.tables.map((t, i) => ({ rid: `rIdT${i + 1}`, idx: i + 1, t }));
    const tablePartsXml = `<tableParts count="${parts.length}">${parts
      .map((p) => `<tablePart r:id="${p.rid}"/>`)
      .join('')}</tableParts>`;
    first.xml = first.xml.replace('</worksheet>', `${tablePartsXml}</worksheet>`);
    const relLines = parts
      .map(
        (p) =>
          `  <Relationship Id="${p.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${p.idx}.xml"/>`,
      )
      .join('\n');
    entries[`xl/worksheets/_rels/${first.fileName}.rels`] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relLines}
</Relationships>`,
    );
    for (const p of parts) {
      entries[`xl/tables/table${p.idx}.xml`] = encoder.encode(buildTableXml(p.idx, p.t));
    }
  }
  if (options.pivotTables && options.pivotTables.length > 0 && sheetParts.length > 0) {
    // A pivot table is referenced ONLY by a worksheet relationship (no element
    // in the sheet XML); merge the rel into any existing sheet rels (E-PIVOT).
    const first = sheetParts[0]!;
    const parts = options.pivotTables.map((p, i) => ({ rid: `rIdP${i + 1}`, idx: i + 1, p }));
    const relLines = parts
      .map(
        (p) =>
          `  <Relationship Id="${p.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable${p.idx}.xml"/>`,
      )
      .join('\n');
    const relsPath = `xl/worksheets/_rels/${first.fileName}.rels`;
    const existing = entries[relsPath] ? decoder.decode(entries[relsPath]) : undefined;
    entries[relsPath] = encoder.encode(
      existing
        ? existing.replace('</Relationships>', `${relLines}\n</Relationships>`)
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relLines}
</Relationships>`,
    );
    for (const p of parts) {
      entries[`xl/pivotTables/pivotTable${p.idx}.xml`] = encoder.encode(
        buildPivotTableXml(p.idx, p.p),
      );
    }
  }
  if (options.slicers && options.slicers.length > 0 && sheetParts.length > 0) {
    const first = sheetParts[0]!;
    const slicerParts = options.slicers.map((s, i) => ({ idx: i + 1, s }));
    // Worksheet rels → slicer parts (merge into sheet1's existing rels).
    const wsRelLines = slicerParts
      .map(
        (p) =>
          `  <Relationship Id="rIdSl${p.idx}" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer${p.idx}.xml"/>`,
      )
      .join('\n');
    const wsRelsPath = `xl/worksheets/_rels/${first.fileName}.rels`;
    const wsExisting = entries[wsRelsPath] ? decoder.decode(entries[wsRelsPath]) : undefined;
    entries[wsRelsPath] = encoder.encode(
      wsExisting
        ? wsExisting.replace('</Relationships>', `${wsRelLines}\n</Relationships>`)
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${wsRelLines}
</Relationships>`,
    );
    // Workbook rels → slicerCache parts.
    const wbRelLines = slicerParts
      .map(
        (p) =>
          `  <Relationship Id="rIdSc${p.idx}" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache${p.idx}.xml"/>`,
      )
      .join('\n');
    const wbRelsPath = 'xl/_rels/workbook.xml.rels';
    entries[wbRelsPath] = encoder.encode(
      decoder
        .decode(entries[wbRelsPath])
        .replace('</Relationships>', `${wbRelLines}\n</Relationships>`),
    );
    // The slicer + slicerCache parts.
    const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    const rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    for (const p of slicerParts) {
      const sl = p.s;
      const styleAttr = sl.styleName ? ` style="${escapeXml(sl.styleName)}"` : '';
      entries[`xl/slicers/slicer${p.idx}.xml`] = encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="${x14}" xmlns:r="${rNs}"><slicer name="${escapeXml(sl.name)}" cache="${escapeXml(
          sl.cacheName,
        )}" caption="${escapeXml(sl.caption)}" columnCount="${sl.columnCount ?? 1}"${styleAttr}/></slicers>`,
      );
      const cache = sl.cache;
      const sourceAttr = cache.sourceName ? ` sourceName="${escapeXml(cache.sourceName)}"` : '';
      const tableExt =
        cache.tableId !== undefined && cache.column !== undefined
          ? `<extLst><ext xmlns:x14="${x14}" uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"><x14:slicerCacheDefinition><x14:tableSlicerCache tableId="${cache.tableId}" column="${cache.column}"/></x14:slicerCacheDefinition></ext></extLst>`
          : '';
      entries[`xl/slicerCaches/slicerCache${p.idx}.xml`] = encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="${x14}" xmlns:r="${rNs}" name="${escapeXml(
          sl.cacheName,
        )}"${sourceAttr}>${tableExt}</slicerCacheDefinition>`,
      );
    }
  }
  for (const s of sheetParts) {
    entries[`xl/worksheets/${s.fileName}`] = encoder.encode(s.xml);
  }
  if (options.stylesXml) {
    entries['xl/styles.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${options.stylesXml}
</styleSheet>`,
    );
  }
  return zipSync(entries);
}

function buildWorkbookXml(
  sheets: ReadonlyArray<XlsxSheetSpec>,
  date1904: boolean,
  definedNames: ReadonlyArray<XlsxDefinedNameSpec>,
): string {
  const lines = sheets.map(
    (s, i) => `    <sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
  );
  const propsXml = date1904 ? '  <workbookPr date1904="1"/>\n' : '';
  const definedNamesXml =
    definedNames.length > 0
      ? '\n  <definedNames>' +
        definedNames
          .map(
            (d) =>
              `<definedName name="${escapeXml(d.name)}"` +
              (d.localSheetId !== undefined ? ` localSheetId="${d.localSheetId}"` : '') +
              `>${escapeXml(d.value)}</definedName>`,
          )
          .join('') +
        '</definedNames>'
      : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${propsXml}  <sheets>
${lines.join('\n')}
  </sheets>${definedNamesXml}
</workbook>`;
}

function buildWorkbookRelsXml(sheetCount: number, includeStyles: boolean): string {
  const lines: Array<string> = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  ];
  for (let i = 0; i < sheetCount; i++) {
    lines.push(
      `  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
    );
  }
  const sharedId = sheetCount + 1;
  lines.push(
    `  <Relationship Id="rId${sharedId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`,
  );
  if (includeStyles) {
    lines.push(
      `  <Relationship Id="rId${sharedId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    );
  }
  lines.push('</Relationships>');
  return lines.join('\n');
}
