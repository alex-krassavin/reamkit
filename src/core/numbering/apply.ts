// Numbering as a FlowDoc transform (ECMA-376 §17.9, ir-design stage 6).
//
// Walks a body, advancing list counters and materializing each numbered
// paragraph's marker as a leading run (plus level-inherited indents). Format
// readers run this while building the FlowDoc, so writers see ready markers;
// the PDF renderer also accepts a raw body + numbering for direct callers.
// Each header/footer band numbers independently (its own NumberingState).

import type {
  BodyElement,
  Numbering,
  NumberingReference,
  Paragraph,
  ParagraphProperties,
  Run,
  StyleSheet,
} from '@/core/document-model';
import { pt } from '@/core/ir';
import { NumberingState, effectiveAbstract } from '@/core/numbering/state';
import { resolveParagraphProperties } from '@/core/style-cascade';

/** The height a picture bullet takes when neither the level nor the paragraph
 * mark states a font size — Word's own default text size. */
const DEFAULT_BULLET_HEIGHT_PT = 11;

/**
 * Apply list numbering to a body as a FlowDoc transform (§17.9): walk the
 * elements (recursing into table cells), advancing list counters and prepending
 * each numbered paragraph's resolved marker as a leading `listMarker` run, with
 * level-inherited indents. Shapes and charts do not advance counters. Returns the
 * body unchanged when there is no numbering.
 *
 * @param body      The body elements to transform.
 * @param numbering The parsed numbering definitions, or `undefined`.
 * @param styles    The style sheet, for a paragraph whose numbering comes from
 *                  its STYLE (§17.9.24) rather than its own `w:numPr`.
 * @returns A new body with markers materialized.
 */
export function applyNumbering(
  body: ReadonlyArray<BodyElement>,
  numbering: Numbering | undefined,
  styles?: StyleSheet,
): Array<BodyElement> {
  if (!numbering || numbering.abstractNums.size === 0) return body.map((b) => b);
  const state = new NumberingState();

  // §17.9.24 — a heading is numbered by its style, not by a `w:numPr` of its
  // own: chtoutline.docx numbers Heading 1 "第 %1 章" that way, and reading the
  // paragraph alone dropped the chapter number from every heading.
  const numberingOf = (p: Paragraph): NumberingReference | undefined =>
    p.properties.numbering ??
    (styles ? resolveParagraphProperties(p.properties, styles).numbering : undefined);

  const transformParagraph = (p: Paragraph): Paragraph => {
    // §17.6.17 — a paragraph whose mark carries the `w:sectPr` and nothing else
    // IS the section break; it is not an item of the list its properties name.
    // Numbered anyway it also took a counter, so section_break_numbering.docx's
    // one real item came out "2." where every reader numbers it 1.
    if (p.properties.sectionBreak === true && p.runs.every((r) => r.text === '')) return p;
    const ref = numberingOf(p);
    if (!ref) return p;
    const marker = state.resolveMarker(numbering, ref);
    if (marker === null) return p;
    const instance = numbering.numInstances.get(ref.numId);
    // §17.9.27 — an instance may redefine a level whole, and the marker's font
    // and indent come from the level it actually numbers by.
    const abstractNum = instance ? effectiveAbstract(numbering, instance) : undefined;
    const level = abstractNum?.levels.get(ref.ilvl);

    // §17.9.9 — a level whose bullet is a PICTURE puts the image where the
    // glyph would go; the `w:lvlText` character beside it is only Word's
    // fallback. FDO74215 draws a bordered square and we drew a dot with the
    // level's double underline running out under the tab.
    const pic = level?.picBullet;
    // §17.9.21 — the size a `w:numPicBullet` states is the PICTURE's, not the
    // bullet's: Word writes 3in for one it draws at the height of the text
    // beside it, and both references draw a small mark whatever the file says.
    // Taken at face value, tdf106606.docx's 235×281pt penguin stood beside every
    // item of its list and turned one page into seven.
    const bulletCapPt =
      level?.runProperties.fontSizePt ??
      p.properties.runProperties?.fontSizePt ??
      DEFAULT_BULLET_HEIGHT_PT;
    const bulletScale = pic ? Math.min(1, bulletCapPt / Math.max(1, pic.heightPt)) : 1;
    const markerRuns: Array<Run> = pic
      ? [
          {
            text: '',
            properties: {},
            listMarker: true,
            inlineImage: {
              resource: pic.resource,
              width: pt(pic.widthPt * bulletScale),
              height: pt(pic.heightPt * bulletScale),
            },
          },
          { text: '\t', properties: {}, listMarker: true },
        ]
      : [
          {
            text: `${marker}\t`,
            // §17.9.6 — the marker is formatted like the PARAGRAPH MARK, with
            // the level's own `w:rPr` on top. Taking the level's alone left the
            // marker at whatever size the paragraph's STYLE lends: fdo78420's
            // bullets are 9pt text under an 11pt list style, and an 11pt bullet
            // made every one of its lines 4pt taller — 41 pages against 23.
            properties: { ...p.properties.runProperties, ...level?.runProperties },
            listMarker: true,
          },
        ];
    const newProps: ParagraphProperties = mergeIndentFromLevel(
      p.properties,
      level?.paragraphProperties,
    );
    return { ...p, properties: newProps, runs: [...markerRuns, ...p.runs] };
  };

  const visit = (el: BodyElement): BodyElement => {
    if (el.kind === 'paragraph') {
      return { kind: 'paragraph', paragraph: transformParagraph(el.paragraph) };
    }
    if (el.kind === 'image') return el;
    // Shapes don't advance list counters (a floating/inline shape isn't a list
    // item). Numbered lists inside a shape's text box are out of scope for M5.
    if (el.kind === 'shape') return el;
    if (el.kind === 'chart') return el;
    return {
      kind: 'table',
      table: {
        ...el.table,
        rows: el.table.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({ ...cell, content: cell.content.map(visit) })),
        })),
      },
    };
  };

  return body.map(visit);
}

function mergeIndentFromLevel(
  paragraphProps: ParagraphProperties,
  levelProps: ParagraphProperties | undefined,
): ParagraphProperties {
  if (!levelProps) return paragraphProps;
  const out: { -readonly [K in keyof ParagraphProperties]: ParagraphProperties[K] } = {
    ...paragraphProps,
  };
  if (out.indentLeft === undefined && levelProps.indentLeft !== undefined) {
    out.indentLeft = levelProps.indentLeft;
  }
  if (out.indentRight === undefined && levelProps.indentRight !== undefined) {
    out.indentRight = levelProps.indentRight;
  }
  if (out.indentFirstLine === undefined && levelProps.indentFirstLine !== undefined) {
    out.indentFirstLine = levelProps.indentFirstLine;
  }
  return out;
}

/**
 * Apply {@link applyNumbering} to each header/footer band independently (each
 * gets its own counter state), keyed as in the input map.
 *
 * @param hf        The header/footer bodies by key, or `undefined`.
 * @param numbering The parsed numbering definitions, or `undefined`.
 * @returns A new map with numbering applied (empty when `hf` is empty/absent).
 */
export function applyNumberingToHeadersFooters(
  hf: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
  numbering: Numbering | undefined,
  styles?: StyleSheet,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  if (!hf || hf.size === 0) return new Map();
  if (!numbering || numbering.abstractNums.size === 0) return hf;
  const out = new Map<string, ReadonlyArray<BodyElement>>();
  for (const [key, value] of hf) {
    out.set(key, applyNumbering(value, numbering, styles));
  }
  return out;
}
