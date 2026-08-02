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
import { NumberingState, effectiveAbstract } from '@/core/numbering/state';
import { resolveParagraphProperties } from '@/core/style-cascade';

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
    const markerRuns: Array<Run> = pic
      ? [
          {
            text: '',
            properties: {},
            listMarker: true,
            inlineImage: { resource: pic.resource, width: pic.widthPt, height: pic.heightPt },
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
