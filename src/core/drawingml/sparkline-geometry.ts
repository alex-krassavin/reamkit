// Sparkline geometry (E-SHEET SC2) — a mini chart for a single cell. Builds
// line / column / win-loss primitives in a local y-up [0,width]×[0,height]
// frame; the layout flips them onto the cell's content rect. A sparkline is a
// bare trend glyph — no axes, legend, or labels (that is what separates it from
// the full chart scene).

import type { StrokeStyle, VectorPath } from '@/core/vector';
import { PathBuilder } from '@/core/vector';

export type SparklineKind = 'line' | 'column' | 'winLoss';

export interface SparklinePrim {
  readonly paths: ReadonlyArray<VectorPath>;
  readonly fillColorHex?: string;
  readonly stroke?: StrokeStyle;
}

// Excel's default sparkline series blue; win/loss "loss" red.
const DEFAULT_SERIES_HEX = '376092';
const LOSS_HEX = 'D00000';

export function buildSparkline(
  kind: SparklineKind,
  values: ReadonlyArray<number>,
  width: number,
  height: number,
  colorHex?: string,
): Array<SparklinePrim> {
  if (values.length === 0 || width <= 0 || height <= 0) return [];
  const color = colorHex ?? DEFAULT_SERIES_HEX;
  switch (kind) {
    case 'line':
      return lineSpark(values, width, height, color);
    case 'column':
      return columnSpark(values, width, height, color);
    case 'winLoss':
      return winLossSpark(values, width, height, color);
  }
}

function lineSpark(
  values: ReadonlyArray<number>,
  w: number,
  h: number,
  color: string,
): Array<SparklinePrim> {
  const n = values.length;
  const pad = Math.min(1, h * 0.12);
  const { min, span } = extent(values);
  const x = (i: number): number => (n === 1 ? w / 2 : (i / (n - 1)) * w);
  const y = (v: number): number => pad + ((v - min) / span) * (h - 2 * pad);
  const b = new PathBuilder().moveTo(x(0), y(values[0]!));
  for (let i = 1; i < n; i++) b.lineTo(x(i), y(values[i]!));
  return [{ paths: [b.build()], stroke: { colorHex: color, widthPt: 0.75, join: 'round' } }];
}

function columnSpark(
  values: ReadonlyArray<number>,
  w: number,
  h: number,
  color: string,
): Array<SparklinePrim> {
  const n = values.length;
  const { min, span } = extent(values);
  const slot = w / n;
  const bw = slot * 0.72;
  const off = (slot - bw) / 2;
  const pad = h * 0.06;
  const paths: Array<VectorPath> = [];
  for (let i = 0; i < n; i++) {
    const frac = (values[i]! - min) / span; // 0 at the lowest value, 1 at the highest
    const bh = Math.max(0.4, pad + frac * (h - pad));
    paths.push(rectPath(i * slot + off, 0, bw, bh));
  }
  return [{ paths, fillColorHex: color }];
}

function winLossSpark(
  values: ReadonlyArray<number>,
  w: number,
  h: number,
  color: string,
): Array<SparklinePrim> {
  const n = values.length;
  const slot = w / n;
  const bw = slot * 0.72;
  const off = (slot - bw) / 2;
  const mid = h / 2;
  const barH = h * 0.4;
  const wins: Array<VectorPath> = [];
  const losses: Array<VectorPath> = [];
  for (let i = 0; i < n; i++) {
    const v = values[i]!;
    if (v === 0) continue;
    const x = i * slot + off;
    if (v > 0) wins.push(rectPath(x, mid, bw, barH));
    else losses.push(rectPath(x, mid - barH, bw, barH));
  }
  const out: Array<SparklinePrim> = [];
  if (wins.length > 0) out.push({ paths: wins, fillColorHex: color });
  if (losses.length > 0) out.push({ paths: losses, fillColorHex: LOSS_HEX });
  return out;
}

// Value extent with a non-zero span (a flat series renders at the baseline).
function extent(values: ReadonlyArray<number>): { min: number; span: number } {
  let min = values[0]!;
  let max = values[0]!;
  for (const v of values) {
    if (v < min) min = v;
    if (v > max) max = v;
  }
  return { min, span: max - min || 1 };
}

function rectPath(x: number, y: number, w: number, h: number): VectorPath {
  return new PathBuilder()
    .moveTo(x, y)
    .lineTo(x + w, y)
    .lineTo(x + w, y + h)
    .lineTo(x, y + h)
    .close()
    .build();
}
