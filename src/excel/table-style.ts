// §18.8.40 — a table's style resolved to the colours of the workbook it is in:
// the one the workbook defines under that name, or the one Excel builds in.

import type {
  Dxf,
  TableStyleElementType,
  TableStyleFormat,
  XlsxBorder,
  XlsxBorderEdge,
  XlsxStyles,
} from '@/core/spreadsheet-model';
import type { PresetColor, PresetDxf, PresetEdge } from '@/excel/table-style-presets';
import type { WorkbookColors } from '@/excel/styles-parser';

import { presetTableStyle } from '@/excel/table-style-presets';
import { workbookColorHex } from '@/excel/styles-parser';

/**
 * The style a table names, region by region, its colours resolved: the
 * workbook's own style of that name when it defines one, else the style Excel
 * builds in (TableStyleLight1…21, Medium1…28, Dark1…11).
 *
 * @param name   The style's name (`<tableStyleInfo name>`).
 * @param styles The workbook's styles: its own table styles and their dxfs.
 * @param colors The workbook's theme and palette, for a built-in style's colours.
 * @returns The style, or undefined for a name that is neither.
 */
export function resolveTableStyleFormat(
  name: string,
  styles: XlsxStyles,
  colors: WorkbookColors,
): TableStyleFormat | undefined {
  const own = styles.tableStyles?.get(name);
  if (own) {
    const out: Partial<Record<TableStyleElementType, { dxf: Dxf; size?: number }>> = {};
    for (const el of own.elements) {
      const dxf = styles.dxfs?.[el.dxfId];
      if (dxf) out[el.type] = { dxf, ...(el.size !== undefined ? { size: el.size } : {}) };
    }
    return out;
  }
  const preset = presetTableStyle(name);
  if (!preset) return undefined;
  const out: Partial<Record<TableStyleElementType, { dxf: Dxf }>> = {};
  for (const [type, dxf] of Object.entries(preset) as Array<[TableStyleElementType, PresetDxf]>) {
    out[type] = { dxf: presetDxf(dxf, colors) };
  }
  return out;
}

/** A preset's format with its theme colours resolved. */
function presetDxf(p: PresetDxf, colors: WorkbookColors): Dxf {
  const hex = (c: PresetColor): string | undefined =>
    workbookColorHex({ '@_theme': String(c[0]), '@_tint': String(c[1]) }, colors);
  const color = p.color ? hex(p.color) : undefined;
  const fill = p.fill ? hex(p.fill) : undefined;
  const edge = (e: PresetEdge | undefined): XlsxBorderEdge | undefined => {
    if (!e) return undefined;
    const c = hex(e[1]);
    return { style: e[0], ...(c !== undefined ? { colorHex: c } : {}) };
  };
  const border: { -readonly [K in keyof XlsxBorder]: XlsxBorder[K] } = {};
  for (const side of ['left', 'right', 'top', 'bottom', 'vertical', 'horizontal'] as const) {
    const e = edge(p.border?.[side]);
    if (e) border[side] = e;
  }
  return {
    ...(p.bold || color !== undefined
      ? {
          font: {
            ...(p.bold ? { bold: true } : {}),
            ...(color !== undefined ? { colorHex: color } : {}),
          },
        }
      : {}),
    ...(fill !== undefined
      ? { fill: { patternType: 'solid', fgColorHex: fill, bgColorHex: fill } }
      : {}),
    ...(Object.keys(border).length > 0 ? { border } : {}),
  };
}
