// The drawings a projected sheet carries, wherever the projection put them: a
// printed grid carries its own over it (TableOverlay), a window and a sheet
// with no grid keep them in the body.

import type { BodyElement } from '@/core/document-model';

/**
 * Every drawing in a projected body, in order — the body's own floats and the
 * ones each printed table carries over it, each once (a drawing across two
 * column bands is in both tables).
 *
 * @param body The projected body.
 * @returns The drawing elements.
 */
export function sheetDrawings(body: ReadonlyArray<BodyElement>): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  const seen = new Set<BodyElement>();
  for (const el of body) {
    const found =
      el.kind === 'table'
        ? (el.table.overlay?.drawings ?? [])
        : el.kind === 'paragraph'
          ? []
          : [el];
    for (const d of found) {
      if (seen.has(d)) continue;
      seen.add(d);
      out.push(d);
    }
  }
  return out;
}
