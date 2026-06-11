// Paragraph-of-segments facade over the UAX #9 algorithm (stage 6 / A5).
//
// Consumers (layout engines) describe a paragraph as a flat list of segments —
// plain text runs and atomic objects (inline images, math boxes) — and get
// back one embedding level per "real" position: one per text code point, one
// per object. The explicit-formatting protocol stays in here: a segment with
// `rtl` is wrapped in RLE…PDF (an LTR segment inside an RTL paragraph in
// LRE…PDF) so run-level direction overrides neutral resolution, and objects
// participate as U+FFFC (Object Replacement Character).

import { computeBidi } from '@/core/bidi/algorithm';

const OBJECT_REPLACEMENT = 0xfffc;
const RLE = 0x202b;
const PDF_FMT = 0x202c;
const LRE = 0x202a;

export interface BidiSegment {
  // Text content; ignored when `object` is true.
  readonly text: string;
  // Atomic neutral object (inline image / math box) — counts as ONE position.
  readonly object?: boolean;
  // Run-level direction override (w:rtl).
  readonly rtl?: boolean;
}

export function segmentLevels(
  segments: ReadonlyArray<BidiSegment>,
  baseDir: 'ltr' | 'rtl',
): Array<number> {
  const bidiCps: Array<number> = [];
  // For each bidi code point, the index of the "real" position it maps to,
  // or -1 for inserted control characters.
  const realIndexOfBidi: Array<number> = [];
  let realCount = 0;

  for (const seg of segments) {
    const isObject = seg.object === true;
    const wrap = seg.rtl ? RLE : baseDir === 'rtl' && !isObject ? LRE : 0;
    if (wrap) {
      bidiCps.push(wrap);
      realIndexOfBidi.push(-1);
    }
    if (isObject) {
      bidiCps.push(OBJECT_REPLACEMENT);
      realIndexOfBidi.push(realCount);
      realCount++;
    } else {
      for (const ch of seg.text) {
        bidiCps.push(ch.codePointAt(0)!);
        realIndexOfBidi.push(realCount);
        realCount++;
      }
    }
    if (wrap) {
      bidiCps.push(PDF_FMT);
      realIndexOfBidi.push(-1);
    }
  }

  const { levels } = computeBidi(bidiCps, baseDir);
  // Project levels back onto real positions.
  const realLevels: Array<number> = new Array<number>(realCount).fill(baseDir === 'rtl' ? 1 : 0);
  for (let i = 0; i < bidiCps.length; i++) {
    const ri = realIndexOfBidi[i]!;
    if (ri >= 0) realLevels[ri] = levels[i]!;
  }
  return realLevels;
}
