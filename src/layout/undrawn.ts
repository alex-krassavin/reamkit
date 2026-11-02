// The characters a laid-out document draws as boxes.
//
// A run is drawn in its own face, and a character that face lacks goes to
// another one: the script face fetched for it, then every family already
// loaded (see `fallbackFaceKey`). When none of them has the character it is
// still drawn, as the face's glyph 0. That is an empty box on the page, and
// nothing in the text layer, because /ToUnicode maps only real glyphs. A line
// of Hindi in a document whose faces were all Latin came out like that, and
// the conversion said nothing: the loss report was empty and `strict` passed.
//
// So the laid-out pages are read back, token by token, against the face each
// token is drawn in. That is the face the emitter shapes it with, so this is
// what the page shows, not what the text asked for.

import type { Loss } from '@/core/ir';
import type { LaidOutPage, Line } from '@/layout/page-doc';

/** How many of the characters a loss names before it only counts the rest. */
const NAMED = 8;

/**
 * The characters the pages draw with no glyph, each once, in the order they
 * first appear.
 *
 * @param pages The laid-out pages.
 * @returns The code points, and the 1-based page each first appears on.
 */
export function undrawnCharacters(pages: ReadonlyArray<LaidOutPage>): Map<number, number> {
  const out = new Map<number, number>();
  let pageNo = 0;
  const readLine = (line: Line): void => {
    for (const t of line.tokens) {
      if (t.kind === 'text') {
        // A tab token's text is the leader that fills it, drawn in the run's
        // face like any other.
        for (const ch of t.text) {
          const cp = ch.codePointAt(0)!;
          if (isControl(cp) || out.has(cp)) continue;
          if (t.font.parsed.glyphForCodepoint(cp) === 0) out.set(cp, pageNo);
        }
      } else if (t.kind === 'math') {
        for (const item of t.items) {
          if (item.kind !== 'glyph') continue;
          for (const ch of item.text) {
            const cp = ch.codePointAt(0)!;
            if (isControl(cp) || out.has(cp)) continue;
            if (item.font.parsed.glyphForCodepoint(cp) === 0) out.set(cp, pageNo);
          }
        }
      } else {
        // MS-EMF / MS-WMF: a metafile sets its own words in lines of its own.
        for (const text of t.metafile?.texts ?? []) readLine(text.line);
      }
    }
  };
  for (const page of pages) {
    pageNo++;
    for (const item of page.commands) if (item.type === 'line') readLine(item.line);
  }
  return out;
}

/**
 * The loss for the characters the pages draw as boxes, or `undefined` when
 * every character has a glyph.
 *
 * @param pages The laid-out pages.
 * @returns One loss that names the characters, at the page the first is on.
 */
export function undrawnLoss(pages: ReadonlyArray<LaidOutPage>): Loss | undefined {
  const undrawn = undrawnCharacters(pages);
  if (undrawn.size === 0) return undefined;
  const cps = [...undrawn.keys()];
  const named = cps
    .slice(0, NAMED)
    .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`)
    .join(', ');
  const rest = cps.length > NAMED ? ` and ${cps.length - NAMED} more` : '';
  return {
    severity: 'dropped',
    feature: 'text',
    detail:
      `no available font has a glyph for ${named}${rest}; ` +
      'they are drawn as empty boxes and are missing from the text layer',
    where: `page ${undrawn.get(cps[0]!)}`,
  };
}

/** C0 and C1 controls: a line break or a tab is never drawn as a glyph. */
function isControl(cp: number): boolean {
  return cp < 0x20 || (cp >= 0x7f && cp < 0xa0);
}
