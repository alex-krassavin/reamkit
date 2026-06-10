// Build a minimal valid xlsx package in memory for tests, optionally with a
// custom styles.xml and merged cell ranges.

import { zipSync } from 'fflate';

import { formatCellRef } from '@/excel';

const encoder = new TextEncoder();

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
  readonly date1904?: boolean;
  readonly definedNames?: ReadonlyArray<XlsxDefinedNameSpec>;
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
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetPrXml}
  ${colsXml}
  <sheetData>
${sheetRows.join('\n')}
  </sheetData>
  ${printOptionsXml}
  ${mergeXml}
  ${marginsXml}
  ${setupXml}
  ${rowBreaksXml}
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
    '\n</Types>';

  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(contentTypes),
    '_rels/.rels': encoder.encode(ROOT_RELS),
    'xl/workbook.xml': encoder.encode(workbookXml),
    'xl/_rels/workbook.xml.rels': encoder.encode(workbookRelsXml),
    'xl/sharedStrings.xml': encoder.encode(sharedStringsXml),
  };
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
