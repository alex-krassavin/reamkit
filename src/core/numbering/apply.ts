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
  Paragraph,
  ParagraphProperties,
  Run,
} from '@/core/document-model';
import { NumberingState } from '@/core/numbering/state';

export function applyNumbering(
  body: ReadonlyArray<BodyElement>,
  numbering: Numbering | undefined,
): Array<BodyElement> {
  if (!numbering || numbering.abstractNums.size === 0) return body.map((b) => b);
  const state = new NumberingState();

  const transformParagraph = (p: Paragraph): Paragraph => {
    const ref = p.properties.numbering;
    if (!ref) return p;
    const marker = state.resolveMarker(numbering, ref);
    if (marker === null) return p;
    const instance = numbering.numInstances.get(ref.numId);
    const abstractNum = instance ? numbering.abstractNums.get(instance.abstractNumId) : undefined;
    const level = abstractNum?.levels.get(ref.ilvl);

    const markerRun: Run = {
      text: `${marker}\t`,
      properties: level?.runProperties ?? {},
    };
    const newProps: ParagraphProperties = mergeIndentFromLevel(
      p.properties,
      level?.paragraphProperties,
    );
    return { properties: newProps, runs: [markerRun, ...p.runs] };
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

export function applyNumberingToHeadersFooters(
  hf: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
  numbering: Numbering | undefined,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  if (!hf || hf.size === 0) return new Map();
  if (!numbering || numbering.abstractNums.size === 0) return hf;
  const out = new Map<string, ReadonlyArray<BodyElement>>();
  for (const [key, value] of hf) {
    out.set(key, applyNumbering(value, numbering));
  }
  return out;
}
