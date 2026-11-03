// Which typefaces a document asks for — read off the parsed model, not the
// bytes of one format.
//
// The substitute downloader needs to know the families before it fetches, and
// it used to learn them from a regex over `w:rFonts` in the .docx XML. That
// left every other format with ONE family for the whole document: a deck whose
// theme is Times and whose body is Calibri was set entirely in the default
// sans, and no per-run resolution could help because only one registry was ever
// built (45541_Header/Footer sit at 0.399 for exactly this reason).
//
// The model already carries the answer, in one vocabulary, for all seven inputs
// — a run's `fontFamily` after the reader has resolved whatever indirection its
// format spells it with (`w:asciiTheme`, `+mn-lt`). So the walk happens here,
// once, and every format gets the same treatment.

import type { BodyElement, RunProperties, ShapeBlock } from '@/core/document-model';
import type { FamilyKey } from '@/core/fonts/remote-fonts';
import type { FontSlot } from '@/core/fonts/scripts';
import type { FlowDoc } from '@/core/ir/flow';
import { resolveFamilyKey } from '@/core/fonts/remote-fonts';
import { fontSlotOf } from '@/core/fonts/scripts';

const SLOTS: ReadonlyArray<FontSlot> = ['ascii', 'hAnsi', 'eastAsia', 'cs'];

/**
 * The curated substitute families a document's text needs, always including the
 * sans default as the fallback for unstyled runs, math and chart labels.
 *
 * §17.3.2.26 — a run names a family per slot, and the text decides which slots
 * are drawn from (see {@link fontSlotOf}): the hAnsi family of a document all
 * in ASCII is never asked for, and is not fetched.
 *
 * @param flow The parsed document.
 * @returns One {@link FamilyKey} per distinct family the document names.
 */
export function familiesInFlow(flow: FlowDoc): Set<FamilyKey> {
  const names: Record<FontSlot, Set<string>> = {
    ascii: new Set(),
    hAnsi: new Set(),
    eastAsia: new Set(),
    cs: new Set(),
  };
  const used = new Set<FontSlot>(['ascii']);
  const add = (props: RunProperties | undefined): void => {
    for (const slot of SLOTS) {
      const name = props?.fontFamily?.[slot];
      if (name !== undefined && name.trim() !== '') names[slot].add(name.trim().toLowerCase());
    }
    if (props?.rtl === true) used.add('cs');
  };
  const read = (text: string): void => {
    for (const ch of text) {
      if (used.size === SLOTS.length) return;
      const cp = ch.codePointAt(0) ?? 0;
      if (cp > 0x7f) used.add(fontSlotOf(cp));
    }
  };

  // The styles come first: a run states a family far less often than it inherits
  // one, and the cascade is resolved after this runs.
  add(flow.styles.defaultRunProperties);
  for (const style of flow.styles.styles.values()) {
    add(style.runProperties);
    add(style.paragraphProperties.runProperties);
    add(style.tableLayer?.runProperties);
    for (const c of style.tableConditions ?? []) add(c.layer.runProperties);
  }
  for (const abstract of flow.numbering?.abstractNums.values() ?? []) {
    // §17.9.6 — a bullet names the symbol face it is drawn in, and that face is
    // a family of the document like any other.
    for (const level of abstract.levels.values()) add(level.runProperties);
  }

  const shapeText = (shape: ShapeBlock): void => {
    if (shape.text) visit(shape.text.content);
    for (const child of shape.children ?? []) shapeText(child.shape);
  };
  const visit = (elements: ReadonlyArray<BodyElement>): void => {
    for (const el of elements) {
      if (el.kind === 'paragraph') {
        add(el.paragraph.properties.runProperties);
        for (const run of el.paragraph.runs) {
          add(run.properties);
          read(run.text);
        }
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
        // …and the drawings over a spreadsheet grid (TableOverlay).
        visit(el.table.overlay?.drawings ?? []);
      } else if (el.kind === 'shape') {
        shapeText(el.shape);
      }
    }
  };
  visit(flow.body);
  for (const band of flow.headersFooters?.values() ?? []) visit(band);
  for (const note of flow.footnotes?.values() ?? []) visit(note);
  for (const note of flow.endnotes?.values() ?? []) visit(note);
  for (const comment of flow.comments?.values() ?? []) visit(comment.content);

  const keys = new Set<FamilyKey>(['arimo']);
  for (const slot of used) for (const name of names[slot]) keys.add(resolveFamilyKey(name));
  return keys;
}
