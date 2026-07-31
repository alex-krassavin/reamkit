// XLSX reader (ir-design §7 + E-SHEET): bytes → SheetDoc → FlowDoc. The reader
// builds the SpreadsheetML IR node (readXlsxToSheetDoc) — a workbook of grid
// sheets with its style table, shared strings, defined names and resolved
// charts — and the print model projects it to flow blocks (projectSheetDoc).
// Document-derived state only; caller conversion options stay with the
// converter/facade.

import type { Chart, DocumentInfo, ShapeBlock } from '@/core/document-model';
import type { CoreProperties, Relationship } from '@/core/opc';
import type { PoNode } from '@/core/po-helpers';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir/loss';
import type {
  Sheet,
  SheetActiveXControl,
  SheetChartRef,
  SheetComment,
  SheetDoc,
  SheetFormControl,
  SheetHyperlink,
  SheetImageRef,
  SheetSlicer,
  SheetSlicerItem,
} from '@/core/ir/sheet';
import type {
  ExcelTable,
  MergedRange,
  ParsedWorksheet,
  PivotTable,
  WorksheetCell,
  XlsxStyles,
} from '@/core/spreadsheet-model';
import type { TableFilterColumn } from '@/excel/table-parser';
import type { VmlShapeBox } from '@/excel/vml-drawing';
import type { SlicerCacheDef, SlicerDef } from '@/excel/slicer-parser';

import type { ProjectSheetOptions } from '@/excel/sheet-to-flow';
import { FEATURES, ResourceStore } from '@/core/ir';
import { OpcPackage, isOoxmlRel, parseCoreProperties } from '@/core/opc';
import {
  EMPTY_XLSX_STYLES,
  parseAreaRef,
  parseSharedStrings,
  parseWorkbook,
  parseWorksheet,
  parseXlsxStyles,
} from '@/excel';
import { bytesInclude, bytesIncludePartName } from '@/core/bytes';
import { parseChart, withChartColorStyle } from '@/core/drawingml/chart-parser';
import { DEFAULT_THEME_PALETTE, makeColorResolver } from '@/core/drawingml/colors';
import {
  parseTheme,
  parseThemeEffectStyles,
  parseThemeFillStyles,
  parseThemeLineWidths,
} from '@/core/drawingml/theme-parser';
import { parseSheetDrawing } from '@/excel/sheet-drawing';
import { parseTablePartFull } from '@/excel/table-parser';
import { parsePivotTablePart } from '@/excel/pivot-table-parser';
import { parseSlicerCachePart, parseSlicerPart } from '@/excel/slicer-parser';
import { parseLegacyComments, parsePersons, parseThreadedComments } from '@/excel/comments-parser';
import { parseFormControlProps } from '@/excel/form-control-parser';
import { parseVmlDrawing } from '@/excel/vml-drawing';
import { parsePrinterSettings } from '@/excel/printer-settings';
import { parseRichValueText } from '@/excel/rich-value';
import {
  activeXBinRelId,
  activeXType,
  activeXTypeFromPart,
  parseActiveX,
  parseActiveXBin,
} from '@/excel/activex-parser';
import { parseSheetShapes } from '@/excel/sheet-shape-parser';

import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { resolveCellText } from '@/excel/print-model';

const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';
const CORE_PROPS_PART = 'docProps/core.xml';
const METADATA_PART = 'xl/metadata.xml';
const RICH_VALUE_STRUCTURE_PART = 'xl/richData/rdrichvaluestructure.xml';
const RICH_VALUE_PART = 'xl/richData/rdrichvalue.xml';
// MS relationship type tails for slicer parts (E-SHEET SV2). The worksheet
// references its slicer parts; the workbook references the slicer caches.
const SLICER_REL_TAIL = '/slicer';
const SLICER_CACHE_REL_TAIL = '/slicerCache';
// Threaded comments + their person directory live in the MS 2017 namespace, so
// they match by relationship-type tail rather than the OOXML transitional helper.
const THREADED_COMMENTS_REL_TAIL = '/threadedComment';
const PERSON_REL_TAIL = '/person';
// A slicer over a huge column is bounded so a crafted file cannot blow up the box.
const MAX_SLICER_ITEMS = 256;

// §18.3.1.19 <control> resolves to either a ctrlProps part (a form control) or
// an activeX#.xml one; the target path is what tells them apart.
const ACTIVEX_PART = /^xl\/activeX\//i;

// A table's location, header depth and value filters — indexed by the table's
// numeric id so a slicer's <tableSlicerCache> can resolve its column's values
// (E-SHEET SV2). Reader-internal; not part of the persisted grid model.
interface TableLoc {
  readonly cells: ReadonlyArray<WorksheetCell>;
  readonly ref: MergedRange;
  readonly headerRows: number;
  readonly filters: ReadonlyArray<TableFilterColumn>;
}

/**
 * Read a `.xlsx` and project it to a {@link FlowDoc} in one step:
 * {@link readXlsxToSheetDoc} then {@link projectSheetDoc}.
 *
 * Parsing itself is lossless — SpreadsheetML maps cleanly onto the sheet IR.
 * The losses come from the projection: the print model's defence-in-depth caps
 * (grid size, per-sheet text budget, sparkline range) clip pathological sheets,
 * and each clip is reported rather than applied in silence.
 *
 * @param xlsx    The `.xlsx` (OPC ZIP) bytes.
 * @param options Projection knobs (the W9 reference date).
 * @returns The flow document plus whatever the projection had to clip.
 */
export function readXlsx(xlsx: Uint8Array, options: ProjectSheetOptions = {}): ReadResult<FlowDoc> {
  // The reader owns the sink so the report cannot be split across callers.
  const losses: Array<Loss> = [];
  const doc = projectSheetDoc(readXlsxToSheetDoc(xlsx), { ...options, losses });
  return { doc, losses };
}

/**
 * Read a `.xlsx` into the {@link SheetDoc} SpreadsheetML IR node. Everything that
 * needs the OPC package — workbook/worksheet/styles/shared-string parsing, chart
 * and drawing resolution, tables, pivots, slicers, hyperlinks, comments and
 * controls — happens here; the SheetDoc → FlowDoc projection (sheet-to-flow) is
 * then a pure transform.
 *
 * @param xlsx The `.xlsx` (OPC ZIP) bytes.
 * @returns The parsed spreadsheet IR tree.
 * @throws Error when `xl/workbook.xml` is missing or the workbook has no sheets.
 */
export function readXlsxToSheetDoc(xlsx: Uint8Array): SheetDoc {
  const pkg = OpcPackage.open(xlsx);
  const workbookData = pkg.getPart(WORKBOOK_PART);
  if (!workbookData) throw new Error('Not a valid xlsx: missing xl/workbook.xml');
  const { sheets, date1904, definedNames } = parseWorkbook(workbookData);
  if (sheets.length === 0) throw new Error('xlsx has no sheets');

  const sharedStringsData = pkg.getPart(SHARED_STRINGS_PART);
  // texts feed every consumer (round-trip, value resolution); runs (W6) carry
  // per-run rich formatting for the strings that have it — render-only.
  const { texts: sharedStrings, runs: sharedStringRuns } = sharedStringsData
    ? parseSharedStrings(sharedStringsData)
    : { texts: [] as ReadonlyArray<string>, runs: [] as ReadonlyArray<undefined> };

  const workbookRels = pkg.getPartRelationships(WORKBOOK_PART);
  // §18.8.3 `<color theme="N">` resolves against the workbook theme, and Excel
  // writes it for anything picked from the standard palette. Without the theme
  // those colours parse to nothing at all: tdf171828.xlsx styles its whole
  // header block `theme="2" tint="-0.5"` and every one of those fills came out
  // unpainted. Same palette the charts and table styles already resolve against.
  const themePalette = buildThemePalette(pkg, workbookRels);

  // A cell can store a legacy error and point at the real value in the rich-value
  // table (§18.3.1.4 `vm`). Workbook-scoped: resolved once, applied per sheet.
  const richValueText = parseRichValueText(
    pkg.getPart(METADATA_PART),
    pkg.getPart(RICH_VALUE_STRUCTURE_PART),
    pkg.getPart(RICH_VALUE_PART),
  );

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseXlsxStyles(stylesData, themePalette) : EMPTY_XLSX_STYLES;
  // Threaded-comment authors (E-SHEET W7): xl/persons/person.xml maps person ids
  // to display names. Workbook-scoped, resolved once and shared across sheets.
  const persons = new Map<string, string>();
  for (const rel of workbookRels) {
    if (!rel.type.endsWith(PERSON_REL_TAIL)) continue;
    const part = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (part) for (const [id, name] of parsePersons(part.data)) persons.set(id, name);
  }
  // Charts keyed by their part path (globally unique across sheets); the
  // theme-backed resolver mirrors the docx reader's so schemeClr references in
  // charts resolve to the workbook's actual accents.
  const chartData = new Map<string, Chart>();
  // Content-addressed store for sheet pictures (W1); populated below as anchors
  // resolve, then handed to the SheetDoc so the renderer can fetch image bytes.
  const resources = new ResourceStore();
  const palette = buildThemePalette(pkg, workbookRels);
  const resolveColor = makeColorResolver(palette);
  // §20.1.4.2.19 `<a:lnRef idx>` indexes the theme's line styles for its width.
  const themeLineWidths = buildThemeLineWidths(pkg, workbookRels);
  const themeFillStyles = buildThemeFillStyles(pkg, workbookRels);
  const themeEffectStyles = buildThemeEffectStyles(pkg, workbookRels);

  const sheetsOut: Array<Sheet> = [];
  // §SV2 slicer-resolution state: tables indexed by id (a slicer's
  // tableSlicerCache may reference a table on another sheet) and the slicer parts
  // found per OUTPUT sheet — both consumed after the loop, once all tables index.
  const tableIndex = new Map<number, TableLoc>();
  const pendingSlicers: Array<{ outIdx: number; defs: ReadonlyArray<SlicerDef> }> = [];
  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    const sheet = sheets[sheetIdx]!;
    const sheetRel = workbookRels.find((r) => r.id === sheet.relationshipId);
    if (!sheetRel) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, sheetRel);
    if (!resolved) continue;
    let worksheet = parseWorksheet(resolved.data);
    worksheet = withPrinterPageSetup(
      worksheet,
      pkg,
      resolved.path,
      pkg.getPartRelationships(resolved.path),
    );
    worksheet = withRichValues(worksheet, richValueText);

    // §20.5: the sheet's drawing part — resolve chart frames, pictures and shapes
    // here; the projection emits a block per frame after the grid (W1 pictures,
    // W2 shapes).
    const charts: Array<SheetChartRef> = [];
    const images: Array<SheetImageRef> = [];
    let shapes: Array<ShapeBlock> | undefined;
    if (worksheet.drawingRelId) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const drawingRel = wsRels.find((r) => r.id === worksheet.drawingRelId);
      const drawing = drawingRel ? pkg.resolveRelatedPart(resolved.path, drawingRel) : undefined;
      if (drawing) {
        const { charts: chartRefs, pictures } = parseSheetDrawing(
          drawing.data,
          drawing.path,
          pkg,
          worksheet,
        );
        for (const ref of chartRefs) {
          if (!chartData.has(ref.chartPartPath)) {
            const chartXml = pkg.getPart(ref.chartPartPath);
            const parsed = chartXml ? parseChart(chartXml, resolveColor) : null;
            if (!parsed) continue;
            chartData.set(
              ref.chartPartPath,
              withChartColorStyle(parsed, pkg, ref.chartPartPath, resolveColor),
            );
          }
          charts.push({
            chartPartPath: ref.chartPartPath,
            widthPt: ref.widthPt,
            heightPt: ref.heightPt,
            xPt: ref.xPt,
            yPt: ref.yPt,
          });
        }
        for (const pic of pictures) {
          const bytes = pkg.getPart(pic.imagePartPath);
          if (!bytes) continue;
          images.push({
            resourceId: resources.put(bytes),
            widthPt: pic.widthPt,
            heightPt: pic.heightPt,
            xPt: pic.xPt,
            yPt: pic.yPt,
          });
        }
        // §20.5.2.30 xdr:sp shapes (W2). The shared DrawingML readers need the
        // preserveOrder PoNode tree, so shapes parse the drawing a second time —
        // gated on a shape open tag (`:sp>`/`:sp `) so chart/picture-only drawings
        // skip it (xdr:spPr / xdr:grpSp do not match).
        if (bytesInclude(drawing.data, ':sp>') || bytesInclude(drawing.data, ':sp ')) {
          const parsed = parseSheetShapes(
            drawing.data,
            worksheet,
            resolveColor,
            themeLineWidths,
            themeFillStyles,
            themeEffectStyles,
          );
          if (parsed.length > 0) shapes = parsed;
        }
      }
    }
    // §18.5: the sheet's table parts — resolve each relationship to its
    // tableN.xml, parse it, and resolve its named style to fill colours against
    // the workbook accent (E-SHEET SC3). The projection applies banded shading.
    let tables: Array<ExcelTable> | undefined;
    if (worksheet.tablePartRelIds && worksheet.tablePartRelIds.length > 0) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const resolvedTables: Array<ExcelTable> = [];
      for (const rid of worksheet.tablePartRelIds) {
        const rel = wsRels.find((r) => r.id === rid);
        const part = rel ? pkg.resolveRelatedPart(resolved.path, rel) : undefined;
        const full = part ? parseTablePartFull(part.data) : undefined;
        if (!full) continue;
        resolvedTables.push(resolveTableStyle(full.table, palette));
        // Index the table by id so a slicer can resolve its column (E-SHEET SV2).
        if (full.id !== undefined) {
          tableIndex.set(full.id, {
            cells: worksheet.cells,
            ref: full.table.ref,
            headerRows: full.table.headerRowCount,
            filters: full.filters,
          });
        }
      }
      if (resolvedTables.length > 0) tables = resolvedTables;
    }

    // §18.3 — the sheet's slicer parts (E-SHEET SV2), resolved after the loop.
    // Record the OUTPUT index this sheet will occupy: continued (skipped) sheets
    // make the output index trail the source index.
    {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const defs: Array<SlicerDef> = [];
      for (const rel of wsRels) {
        if (!rel.type.endsWith(SLICER_REL_TAIL)) continue;
        const part = pkg.resolveRelatedPart(resolved.path, rel);
        if (part) defs.push(...parseSlicerPart(part.data));
      }
      if (defs.length > 0) pendingSlicers.push({ outIdx: sheetsOut.length, defs });
    }

    // §18.10: the sheet's pivot tables — referenced ONLY via the worksheet's
    // relationships (there is no element in the sheet XML), so enumerate the rels
    // by type. The output cells are already cached in the grid; PV1 just records
    // the location + named style for PV2 to band (E-PIVOT).
    let pivotTables: Array<PivotTable> | undefined;
    {
      const resolvedPivots: Array<PivotTable> = [];
      for (const rel of pkg.getPartRelationships(resolved.path)) {
        if (!isOoxmlRel(rel.type, 'pivotTable')) continue;
        const part = pkg.resolveRelatedPart(resolved.path, rel);
        const parsed = part ? parsePivotTablePart(part.data) : undefined;
        if (parsed) resolvedPivots.push(resolvePivotStyle(parsed, palette));
      }
      if (resolvedPivots.length > 0) pivotTables = resolvedPivots;
    }

    // §18.3.1.47 cell hyperlinks (W3): resolve each relId to its external URL via
    // the worksheet rels; location-only (in-workbook) links carry no URL.
    let hyperlinks: Array<SheetHyperlink> | undefined;
    if (worksheet.hyperlinks && worksheet.hyperlinks.length > 0) {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const resolvedLinks: Array<SheetHyperlink> = [];
      for (const h of worksheet.hyperlinks) {
        if (h.relId === undefined) continue; // in-workbook location link → no URL
        const rel = wsRels.find((r) => r.id === h.relId);
        if (!rel || rel.targetMode !== 'External' || !rel.target) continue;
        const area = parseAreaRef(h.ref);
        if (!area) continue;
        resolvedLinks.push({
          ref: {
            startColumn: area.startColumn,
            startRow: area.startRow,
            endColumn: area.endColumn,
            endRow: area.endRow,
          },
          url: rel.target,
        });
      }
      if (resolvedLinks.length > 0) hyperlinks = resolvedLinks;
    }

    // §18.7 cell comments / notes (W7): legacy xl/comments and modern threaded
    // comments, both resolved through the worksheet rels. Legacy notes precede the
    // threaded conversation so a cell that has both reads oldest-first.
    let comments: Array<SheetComment> | undefined;
    {
      const wsRels = pkg.getPartRelationships(resolved.path);
      const resolvedComments: Array<SheetComment> = [];
      for (const rel of wsRels) {
        if (!isOoxmlRel(rel.type, 'comments')) continue;
        const part = pkg.resolveRelatedPart(resolved.path, rel);
        if (part) resolvedComments.push(...parseLegacyComments(part.data));
      }
      for (const rel of wsRels) {
        if (!rel.type.endsWith(THREADED_COMMENTS_REL_TAIL)) continue;
        const part = pkg.resolveRelatedPart(resolved.path, rel);
        if (part) resolvedComments.push(...parseThreadedComments(part.data, persons));
      }
      if (resolvedComments.length > 0) comments = resolvedComments;
    }

    // §18.3.* form controls (W8): resolve each control's relId to its ctrlProp
    // part (objectType + state). The projection lists them after the grid.
    const wsRels = pkg.getPartRelationships(resolved.path);
    const resolvedAx: Array<SheetActiveXControl> = [];

    // The visible state of an ActiveX control: its property bag, plus — for a
    // persistStreamInit control, which keeps no <ax:ocxPr> — the caption /
    // value / group name recovered from the binary activeX#.bin (resolved
    // through the activeX#.xml part's own relationships). The property bag wins
    // where both carry a value, so a normal property-bag control is unchanged.
    const activeXState = (part: { path: string; data: Uint8Array }): SheetActiveXControl => {
      const props = parseActiveX(part.data);
      const binRelId = activeXBinRelId(part.data);
      const binRel = binRelId
        ? pkg.getPartRelationships(part.path).find((r) => r.id === binRelId)
        : undefined;
      const binPart = binRel ? pkg.resolveRelatedPart(part.path, binRel) : undefined;
      const binProps = binPart ? parseActiveXBin(binPart.data) : {};
      return { type: 'control', ...binProps, ...props };
    };

    // The legacy VML drawing is read before the <control> list because it is
    // the only part that carries geometry — an ActiveX control's box lives in
    // the `Pict` shape that shares its `shapeId`, nowhere else.
    const legacyVml = readLegacyVml(pkg, resolved.path, wsRels, worksheet);

    let formControls: Array<SheetFormControl> | undefined;
    if (worksheet.formControls && worksheet.formControls.length > 0) {
      const resolvedControls: Array<SheetFormControl> = [];
      for (const fc of worksheet.formControls) {
        // §18.3.1.20 `<controlPr print="0">`, and the same thing said the
        // legacy way by the VML shape this control points at. Excel's "Print
        // object" is on by default; a control that clears it is on screen only,
        // and button-form-control.xlsx — which says it BOTH ways — prints as a
        // blank page in Calc while we drew the button.
        if (fc.print === false) continue;
        if (fc.shapeId !== undefined && legacyVml.nonPrinting.has(fc.shapeId)) continue;
        const rel = wsRels.find((r) => r.id === fc.relId);
        const part = rel ? pkg.resolveRelatedPart(resolved.path, rel) : undefined;
        // §18.3.1.19 <control> reaches BOTH kinds: a form control's ctrlProps
        // part and an ActiveX control's activeX#.xml. Only the relationship
        // target says which, and reading an ocx part as a ctrlProps one yields
        // nothing — which is how a sheet of option buttons came out as a list
        // of bare names, its captions and values sitting unread in the .bin.
        // The <control> element carries no progId, so the type comes from the
        // class id.
        const box = fc.shapeId !== undefined ? legacyVml.boxes.get(fc.shapeId) : undefined;
        if (part && ACTIVEX_PART.test(part.path)) {
          resolvedAx.push({
            ...activeXState(part),
            type: activeXTypeFromPart(part.data),
            ...(fc.name ? { name: fc.name } : {}),
            ...(box ? { box } : {}),
          });
          continue;
        }
        const props = part ? parseFormControlProps(part.data) : {};
        resolvedControls.push({
          ...(fc.name ? { name: fc.name } : {}),
          ...props,
          ...(box ? { box } : {}),
        });
      }
      if (resolvedControls.length > 0) formControls = resolvedControls;
    }

    // A control put on the sheet by Excel's Forms toolbar is declared ONLY in
    // the legacy VML drawing — no `<control>` entry, no ctrlProps part. Reading
    // just the `<controls>` list showed tdf111980_radioButtons.xlsx's five
    // ActiveX buttons and silently lost the five form radio buttons and the
    // group box beside them. Shapes whose id matches a `<control shapeId>` are
    // the ActiveX ones, already resolved above.
    if (legacyVml.controls.length > 0) {
      formControls = [...(formControls ?? []), ...legacyVml.controls];
    }

    // §18.3.* ActiveX / OLE controls (W10): resolve each oleObject's relId to its
    // activeX part (progId → type + the property bag's visible state). Listed
    // after the grid, like form controls.
    if (worksheet.oleObjects && worksheet.oleObjects.length > 0) {
      for (const ole of worksheet.oleObjects) {
        const rel = wsRels.find((r) => r.id === ole.relId);
        const part = rel ? pkg.resolveRelatedPart(resolved.path, rel) : undefined;
        resolvedAx.push({
          ...(part ? activeXState(part) : {}),
          type: activeXType(ole.progId),
        });
      }
    }
    const activeXControls = resolvedAx.length > 0 ? resolvedAx : undefined;

    const grid =
      tables || pivotTables
        ? { ...worksheet, ...(tables ? { tables } : {}), ...(pivotTables ? { pivotTables } : {}) }
        : worksheet;
    sheetsOut.push({
      name: sheet.name,
      ...(sheet.hidden ? { hidden: true } : {}),
      grid,
      ...(charts.length > 0 ? { charts } : {}),
      ...(images.length > 0 ? { images } : {}),
      ...(shapes ? { shapes } : {}),
      ...(hyperlinks ? { hyperlinks } : {}),
      ...(comments ? { comments } : {}),
      ...(formControls ? { formControls } : {}),
      ...(activeXControls ? { activeXControls } : {}),
    });
  }

  // §SV2 — resolve slicer panels now that every table is indexed. Slicer caches
  // are workbook-scoped (referenced from the workbook rels); each binds by name
  // to a slicer's @cache. A panel renders after its sheet's grid + charts.
  if (pendingSlicers.length > 0) {
    const cacheByName = new Map<string, SlicerCacheDef>();
    for (const rel of workbookRels) {
      if (!rel.type.endsWith(SLICER_CACHE_REL_TAIL)) continue;
      const part = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
      const cache = part ? parseSlicerCachePart(part.data) : undefined;
      if (cache) cacheByName.set(cache.name, cache);
    }
    for (const { outIdx, defs } of pendingSlicers) {
      const sheet = sheetsOut[outIdx];
      if (!sheet) continue;
      const slicers = defs.map((def) =>
        resolveSlicer(
          def,
          cacheByName.get(def.cacheName),
          tableIndex,
          styles,
          sharedStrings,
          date1904,
          palette,
        ),
      );
      sheetsOut[outIdx] = { ...sheet, slicers };
    }
  }

  // §21.2.2.59/.215 — a chart written without caches reads its numbers, its
  // categories and its series names from the workbook itself. 123233_charts
  // .xlsx does exactly that, and taken as cache-only its four charts drew as
  // empty axes with a legend of "Series 1"… Resolved once here, where every
  // sheet is finally in hand: a chart on sheet 1 routinely points at sheet 4.
  for (const [path, chart] of chartData) {
    const resolved = withWorkbookData(chart, sheetsOut, styles, sharedStrings, date1904);
    if (resolved !== chart) chartData.set(path, resolved);
  }

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;
  const info = infoFromCore(coreProps);

  return {
    kind: 'sheet',
    sheets: sheetsOut,
    styles,
    sharedStrings,
    // Carry rich runs only when some shared string actually has them (W6) — keeps
    // the common plain-text workbook's SheetDoc unchanged.
    ...(sharedStringRuns.some((r) => r !== undefined) ? { sharedStringRuns } : {}),
    definedNames,
    date1904,
    ...(chartData.size > 0 ? { chartData } : {}),
    resources,
    ...(info ? { info } : {}),
  };
}

// Theme palette: the workbook's theme part merged over the built-in Office
// defaults (the docx reader's pattern). Drives both chart schemeClr resolution
// and the table-style accent (E-SHEET SC3).
/**
 * Fill in the paper size and orientation a `<pageSetup>` leaves to its
 * `printerSettings` part.
 *
 * `<pageSetup>` naming no `paperSize` does not mean "the default": Excel
 * records the print dialog's choice in the DEVMODE of the related part, and
 * LibreOffice reads it. simple-monthly-budget.xlsx and 45540_classic_Header.xlsx
 * both print on Letter that way while we assumed A4 — the most visible
 * difference on every page of either. What the sheet states itself always wins.
 */
/**
 * Replace the legacy stand-in in every cell that points at a rich value we can
 * read. Spill.xlsx caches `#VALUE!` in three cells whose rich value says
 * `#SPILL!` — the error Excel shows and the one that explains the sheet, since
 * the file also records which cell three rows down does the blocking.
 *
 * @param worksheet The parsed worksheet.
 * @param richValueText `vm` index → text, from {@link parseRichValueText}.
 * @returns The worksheet, unchanged when nothing resolved (byte-identical path).
 */

/**
 * A chart with its uncached references resolved against the workbook.
 *
 * Only what is missing is filled in: a series that cached its values keeps
 * them, and a reference naming a sheet this workbook does not have resolves to
 * nothing rather than to zeros.
 *
 * @param chart         The parsed chart.
 * @param sheets        Every sheet in the workbook, by tab order.
 * @param styles        The style table (for a referenced cell's number format).
 * @param sharedStrings The shared-string table.
 * @param date1904      The workbook's date system.
 * @returns The chart, unchanged when nothing needed resolving.
 */
function withWorkbookData(
  chart: Chart,
  sheets: ReadonlyArray<Sheet>,
  styles: XlsxStyles,
  sharedStrings: ReadonlyArray<string>,
  date1904: boolean,
): Chart {
  const cellsOf = (ref: string | undefined): Array<string> | undefined => {
    if (ref === undefined) return undefined;
    const area = resolveChartRef(ref, sheets);
    if (!area) return undefined;
    const out: Array<string> = [];
    for (let row = area.startRow; row <= area.endRow; row++) {
      for (let col = area.startColumn; col <= area.endColumn; col++) {
        const cell = area.grid.cells.find((c) => c.row === row && c.column === col);
        out.push(cell ? resolveCellText(cell, sharedStrings, styles, date1904) : '');
      }
    }
    return out;
  };

  let changed = false;
  const series = chart.series.map((s) => {
    if (s.values.length > 0 && s.name !== undefined) return s;
    const values =
      s.values.length > 0 ? s.values : cellsOf(s.valuesRef)?.map((t) => Number(t) || 0);
    const name = s.name ?? cellsOf(s.nameRef)?.find((t) => t.length > 0);
    if (!values && name === undefined) return s;
    changed = true;
    return {
      ...s,
      ...(values ? { values } : {}),
      ...(name !== undefined ? { name } : {}),
    };
  });
  const categories = chart.categories.length > 0 ? chart.categories : cellsOf(chart.categoriesRef);
  if (categories && categories.length > 0 && chart.categories.length === 0) changed = true;
  if (!changed) return chart;
  return {
    ...chart,
    series,
    ...(categories && categories.length > 0 ? { categories } : {}),
  };
}

/** `Sheet!A1:B2` → the sheet's grid and the 0-based rectangle it names. */
function resolveChartRef(
  ref: string,
  sheets: ReadonlyArray<Sheet>,
):
  | {
      grid: ParsedWorksheet;
      startRow: number;
      endRow: number;
      startColumn: number;
      endColumn: number;
    }
  | undefined {
  const bang = ref.lastIndexOf('!');
  if (bang < 0) return undefined;
  const name = ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
  const sheet = sheets.find((s) => s.name === name);
  if (!sheet) return undefined;
  const area = parseAreaRef(ref.slice(bang + 1).replace(/\$/g, ''));
  return area ? { grid: sheet.grid, ...area } : undefined;
}

function withRichValues(
  worksheet: ParsedWorksheet,
  richValueText: ReadonlyMap<number, string>,
): ParsedWorksheet {
  if (richValueText.size === 0) return worksheet;
  const resolves = (cell: WorksheetCell): string | undefined =>
    cell.valueMetadataIndex !== undefined ? richValueText.get(cell.valueMetadataIndex) : undefined;
  // A workbook can carry rich values none of THIS sheet's cells point at; leave
  // it the object it was so nothing downstream sees a change that is not one.
  if (!worksheet.cells.some((cell) => resolves(cell) !== undefined)) return worksheet;
  const cells = worksheet.cells.map((cell) => {
    const text = resolves(cell);
    return text === undefined ? cell : { ...cell, rawValue: text };
  });
  return { ...worksheet, cells };
}

function withPrinterPageSetup(
  worksheet: ParsedWorksheet,
  pkg: OpcPackage,
  sheetPath: string,
  wsRels: ReadonlyArray<Relationship>,
): ParsedWorksheet {
  const setup = worksheet.pageSetup;
  const relId = setup?.printerSettingsRelId;
  if (!setup || relId === undefined) return worksheet;
  // The id has done its job once resolved, and it must not survive into the
  // model: a relationship id is a spelling, not a fact about the document, and
  // two dialects of the same workbook name the same part differently.
  const { printerSettingsRelId: _resolved, ...rest } = setup;
  void _resolved;
  const rel = wsRels.find((r) => r.id === relId);
  const part = rel ? pkg.resolveRelatedPart(sheetPath, rel) : undefined;
  const printer = part ? parsePrinterSettings(part.data) : {};
  return {
    ...worksheet,
    pageSetup: {
      ...rest,
      ...(rest.paperSize === undefined && printer.paperSize !== undefined
        ? { paperSize: printer.paperSize }
        : {}),
      ...(rest.orientation === undefined && printer.orientation !== undefined
        ? { orientation: printer.orientation }
        : {}),
    },
  };
}

/**
 * The sheet's form controls that live only in its legacy VML drawing (§18.3.1.36).
 *
 * Shapes that back an ActiveX control are skipped — their `o:spid` is the
 * `shapeId` of a `<control>` element, which the caller has already resolved
 * through the activeX part. What is left is the Forms-toolbar kind, whose
 * caption and checked state exist nowhere else in the package.
 */
function readLegacyVml(
  pkg: OpcPackage,
  sheetPath: string,
  wsRels: ReadonlyArray<Relationship>,
  worksheet: ParsedWorksheet,
): {
  controls: Array<SheetFormControl>;
  boxes: ReadonlyMap<string, VmlShapeBox>;
  nonPrinting: ReadonlySet<string>;
} {
  const empty = {
    controls: [],
    boxes: new Map<string, VmlShapeBox>(),
    nonPrinting: new Set<string>(),
  };
  const relId = worksheet.legacyDrawingRelId;
  if (relId === undefined) return empty;
  const rel = wsRels.find((r) => r.id === relId);
  const part = rel ? pkg.resolveRelatedPart(sheetPath, rel) : undefined;
  if (!part) return empty;
  const activeXShapeIds = new Set(
    (worksheet.formControls ?? []).map((fc) => fc.shapeId).filter((id) => id !== undefined),
  );
  const drawing = parseVmlDrawing(part.data);
  const out: Array<SheetFormControl> = [];
  for (const shape of drawing.controls) {
    if (shape.shapeId !== undefined && activeXShapeIds.has(shape.shapeId)) continue;
    out.push({
      objectType: shape.objectType,
      ...(shape.caption ? { name: shape.caption, caption: shape.caption } : {}),
      ...(shape.checked ? { checked: true } : {}),
      ...(shape.box ? { box: shape.box } : {}),
      ...(shape.fontSizePt !== undefined ? { fontSizePt: shape.fontSizePt } : {}),
    });
  }
  return { controls: out, boxes: drawing.boxes, nonPrinting: drawing.nonPrinting };
}

function buildThemePalette(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): Map<string, string> {
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const rel of workbookRels) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (!resolved) continue;
    for (const [slot, hex] of parseTheme(resolved.data)) palette.set(slot, hex);
    break;
  }
  return palette;
}

function buildThemeLineWidths(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): Array<number> {
  for (const rel of workbookRels) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (resolved) return parseThemeLineWidths(resolved.data);
  }
  return [];
}

// §20.1.4.1.15 — the effect styles an `<a:effectRef idx>` indexes into.
function buildThemeEffectStyles(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): Array<PoNode> {
  for (const rel of workbookRels) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (resolved) return parseThemeEffectStyles(resolved.data);
  }
  return [];
}

// §20.1.4.1.13 — the fill styles an `<a:fillRef idx>` indexes into.
function buildThemeFillStyles(
  pkg: OpcPackage,
  workbookRels: ReadonlyArray<Relationship>,
): Array<PoNode> {
  for (const rel of workbookRels) {
    if (!isOoxmlRel(rel.type, 'theme')) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, rel);
    if (resolved) return parseThemeFillStyles(resolved.data);
  }
  return [];
}

// Resolve a table's named built-in style to header / band fill colours against
// the workbook theme. The definitions live in Excel (not the file), but the
// name encodes the accent: TableStyle{Light|Medium|Dark}{N}, where the gallery
// is 7 columns wide — column (N-1)%7 picks the accent (0 = the neutral/grey
// column, 1..6 = accent1..6). Medium/Dark take a solid accent header with white
// text; Light a tinted header with black text; the band is a light tint. A
// style-less / unrecognized / TableStyleNone table is left uncoloured.
function resolveTableStyle(t: ExcelTable, palette: ReadonlyMap<string, string>): ExcelTable {
  const m = t.styleName ? /TableStyle(Light|Medium|Dark)(\d+)/i.exec(t.styleName) : null;
  if (!m) return t;
  const kind = m[1]!.toLowerCase();
  const column = (Number(m[2]) - 1) % 7;
  const base = column === 0 ? '7F7F7F' : (palette.get(`accent${column}`) ?? '4472C4');
  if (kind === 'light') {
    return { ...t, headerHex: lighten(base, 0.6), bandHex: lighten(base, 0.85) };
  }
  // medium / dark: a solid accent header with white text.
  return { ...t, headerHex: base, bandHex: lighten(base, 0.8), headerTextHex: 'FFFFFF' };
}

// Resolve a pivot's named built-in style to header / band colours. Pivot styles
// (PivotStyle{Light|Medium|Dark}{N}) live in Excel, not the file; we approximate
// with the same accent-column heuristic as table styles — the pivot gallery
// differs in exact numbering, refined later (E-PIVOT PV2). A style-less /
// unrecognized pivot is left uncoloured (it then renders as a plain grid).
function resolvePivotStyle(p: PivotTable, palette: ReadonlyMap<string, string>): PivotTable {
  const m = p.styleName ? /PivotStyle(Light|Medium|Dark)(\d+)/i.exec(p.styleName) : null;
  if (!m) return p;
  const kind = m[1]!.toLowerCase();
  const column = (Number(m[2]) - 1) % 7;
  const base = column === 0 ? '7F7F7F' : (palette.get(`accent${column}`) ?? '4472C4');
  if (kind === 'light') {
    return { ...p, headerHex: lighten(base, 0.6), bandHex: lighten(base, 0.85) };
  }
  // medium / dark: a solid accent header with white text.
  return { ...p, headerHex: base, bandHex: lighten(base, 0.8), headerTextHex: 'FFFFFF' };
}

// §SV2 — resolve a slicer definition + its cache into a renderable panel: the
// caption, the column-button items with selection, and the style accent. A
// native-table cache resolves items from the referenced table column; an
// OLAP/pivot cache (no <tableSlicerCache>) yields a caption-only box.
function resolveSlicer(
  def: SlicerDef,
  cache: SlicerCacheDef | undefined,
  tableIndex: ReadonlyMap<number, TableLoc>,
  styles: XlsxStyles,
  sharedStrings: ReadonlyArray<string>,
  date1904: boolean,
  palette: ReadonlyMap<string, string>,
): SheetSlicer {
  const items = cache
    ? resolveTableSlicerItems(cache, tableIndex, styles, sharedStrings, date1904)
    : [];
  return {
    caption: def.caption,
    columnCount: def.columnCount,
    items,
    ...resolveSlicerStyle(def.styleName, palette),
  };
}

// The distinct values of the slicer's table column become its buttons (first-seen
// row order, header rows skipped, bounded by MAX_SLICER_ITEMS). Selection comes
// from the column's autofilter when present; an unfiltered column shows
// everything as selected (a freshly-saved slicer's default).
function resolveTableSlicerItems(
  cache: SlicerCacheDef,
  tableIndex: ReadonlyMap<number, TableLoc>,
  styles: XlsxStyles,
  sharedStrings: ReadonlyArray<string>,
  date1904: boolean,
): Array<SheetSlicerItem> {
  const loc = cache.tableId !== undefined ? tableIndex.get(cache.tableId) : undefined;
  if (!loc || cache.columnId === undefined) return [];
  const colOffset = cache.columnId - 1; // table column id (1-based) → range offset
  const absCol = loc.ref.startColumn + colOffset;
  if (absCol < loc.ref.startColumn || absCol > loc.ref.endColumn) return [];
  const firstDataRow = loc.ref.startRow + Math.max(1, loc.headerRows);
  const filter = loc.filters.find((f) => f.colId === colOffset);
  const kept = filter ? new Set(filter.values) : undefined;
  // The column's data cells in row order, resolved to display text.
  const byRow = new Map<number, string>();
  for (const c of loc.cells) {
    if (c.column !== absCol || c.row < firstDataRow || c.row > loc.ref.endRow) continue;
    byRow.set(c.row, resolveCellText(c, sharedStrings, styles, date1904));
  }
  const seen = new Set<string>();
  const out: Array<SheetSlicerItem> = [];
  for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
    const label = byRow.get(row)!;
    if (label === '' || seen.has(label)) continue;
    seen.add(label);
    out.push({ label, selected: kept ? kept.has(label) : true });
    if (out.length >= MAX_SLICER_ITEMS) break;
  }
  return out;
}

// SlicerStyle{Light|Dark|Other}{N} → an accent for the header + selected buttons
// (the same accent-column heuristic as table/pivot styles; the slicer gallery's
// exact numbering differs, refined later). White text reads on the accent; a
// style-less slicer falls back to accent1.
function resolveSlicerStyle(
  styleName: string | undefined,
  palette: ReadonlyMap<string, string>,
): { headerHex: string; headerTextHex: string; selectedHex: string; selectedTextHex: string } {
  const m = styleName ? /SlicerStyle[A-Za-z]*?(\d+)/i.exec(styleName) : null;
  const column = m ? (Number(m[1]) - 1) % 7 : 1;
  const base = column === 0 ? '7F7F7F' : (palette.get(`accent${column}`) ?? '4472C4');
  return { headerHex: base, headerTextHex: 'FFFFFF', selectedHex: base, selectedTextHex: 'FFFFFF' };
}

// Lighten a 6-hex colour toward white by `amount` (0..1).
function lighten(hex: string, amount: number): string {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex);
  const n = m ? parseInt(m[1]!, 16) : 0x4472c4;
  const ch = (shift: number): number => {
    const c = (n >> shift) & 0xff;
    return Math.round(c + (255 - c) * amount);
  };
  const hx = (c: number): string => c.toString(16).padStart(2, '0').toUpperCase();
  return `${hx(ch(16))}${hx(ch(8))}${hx(ch(0))}`;
}

/**
 * The OOXML `.xlsx` {@link DocumentReader}: sniffs the OPC ZIP for
 * `xl/workbook.xml` and reads it into a {@link SheetDoc}. The reader's native
 * tree is the SheetDoc; the facade/Ream project it to a {@link FlowDoc} for
 * rendering (E-SHEET SB1).
 */
export const xlsxReader: DocumentReader<SheetDoc> = {
  id: 'xlsx',
  // The reader's native tree is the SheetDoc; the facade/Ream project it to a
  // FlowDoc for rendering (E-SHEET SB1).
  produces: 'sheet',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && bytesIncludePartName(bytes, 'xl/workbook.xml'),
  read: (bytes) => ({ doc: readXlsxToSheetDoc(bytes), losses: [] }),
};

function infoFromCore(core: CoreProperties | undefined): DocumentInfo | undefined {
  if (!core) return undefined;
  return {
    ...(core.title ? { title: core.title } : {}),
    ...(core.creator ? { author: core.creator } : {}),
    ...(core.subject ? { subject: core.subject } : {}),
    ...(core.keywords ? { keywords: core.keywords } : {}),
    ...(core.created ? { creationDate: core.created } : {}),
    ...(core.modified ? { modificationDate: core.modified } : {}),
  };
}
