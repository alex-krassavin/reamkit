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
// — a run's `fontFamily.ascii` after the reader has resolved whatever
// indirection its format spells it with (`w:asciiTheme`, `+mn-lt`). So the walk
// happens here, once, and every format gets the same treatment.

import type { BodyElement, ShapeBlock } from '@/core/document-model';
import type { FamilyKey } from '@/core/fonts/remote-fonts';
import type { FlowDoc } from '@/core/ir/flow';
import { resolveFamilyKey } from '@/core/fonts/remote-fonts';

/**
 * The curated substitute families a document's text needs, always including the
 * sans default as the fallback for unstyled runs, math and chart labels.
 *
 * @param flow The parsed document.
 * @returns One {@link FamilyKey} per distinct family the document names.
 */
export function familiesInFlow(flow: FlowDoc): Set<FamilyKey> {
  const names = new Set<string>();
  const add = (name: string | undefined): void => {
    if (name !== undefined && name.trim() !== '') names.add(name.trim().toLowerCase());
  };

  // The styles come first: a run states a family far less often than it inherits
  // one, and the cascade is resolved after this runs.
  add(flow.styles.defaultRunProperties.fontFamily?.ascii);
  for (const style of flow.styles.styles.values()) {
    add(style.runProperties.fontFamily?.ascii);
    add(style.paragraphProperties.runProperties?.fontFamily?.ascii);
    add(style.tableLayer?.runProperties?.fontFamily?.ascii);
    for (const c of style.tableConditions ?? []) add(c.layer.runProperties?.fontFamily?.ascii);
  }
  for (const abstract of flow.numbering?.abstractNums.values() ?? []) {
    // §17.9.6 — a bullet names the symbol face it is drawn in, and that face is
    // a family of the document like any other.
    for (const level of abstract.levels.values()) add(level.runProperties.fontFamily?.ascii);
  }

  const shapeText = (shape: ShapeBlock): void => {
    if (shape.text) visit(shape.text.content);
    for (const child of shape.children ?? []) shapeText(child.shape);
  };
  const visit = (elements: ReadonlyArray<BodyElement>): void => {
    for (const el of elements) {
      if (el.kind === 'paragraph') {
        add(el.paragraph.properties.runProperties?.fontFamily?.ascii);
        for (const run of el.paragraph.runs) add(run.properties.fontFamily?.ascii);
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
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
  for (const name of names) keys.add(resolveFamilyKey(name));
  return keys;
}
