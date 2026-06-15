// Build a minimal valid PresentationML pptx in memory for tests (E-PPTX). Emits
// only what the reader needs: [Content_Types].xml, _rels/.rels,
// ppt/presentation.xml (slide size + slide order) and one ppt/slides/slideN.xml
// per slide, with the matching part relationships. Slide layouts / masters /
// theme are omitted — the reader treats them as optional (PX2 adds them).

import { zipSync } from 'fflate';

const encoder = new TextEncoder();

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

export interface BuildPptxOptions {
  /** Slide size in EMU (default a 16:9 deck, 13.333" × 7.5"). */
  readonly cx?: number;
  readonly cy?: number;
}

// `slides[i]` is the inner XML of that slide's `<p:spTree>` (default empty).
export function buildPptx(
  slides: ReadonlyArray<string>,
  options: BuildPptxOptions = {},
): Uint8Array {
  const cx = options.cx ?? 12192000;
  const cy = options.cy ?? 6858000;
  const n = Math.max(1, slides.length);

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Types xmlns="${PKG_REL_NS.replace('relationships', 'content-types')}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    Array.from(
      { length: n },
      (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${SLIDE_CT}"/>`,
    ).join('') +
    `</Types>`;

  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="ppt/presentation.xml"/>` +
    `</Relationships>`;

  const presentation =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<p:presentation xmlns:p="${P_NS}" xmlns:r="${R_NS}">` +
    `<p:sldIdLst>` +
    Array.from({ length: n }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join('') +
    `</p:sldIdLst>` +
    `<p:sldSz cx="${cx}" cy="${cy}"/>` +
    `</p:presentation>`;

  const presRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    Array.from(
      { length: n },
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="${R_NS}/slide" Target="slides/slide${i + 1}.xml"/>`,
    ).join('') +
    `</Relationships>`;

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(contentTypes),
    '_rels/.rels': encoder.encode(rootRels),
    'ppt/presentation.xml': encoder.encode(presentation),
    'ppt/_rels/presentation.xml.rels': encoder.encode(presRels),
  };
  for (let i = 0; i < n; i++) {
    const slide =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
      `<p:cSld><p:spTree>${slides[i] ?? ''}</p:spTree></p:cSld>` +
      `</p:sld>`;
    files[`ppt/slides/slide${i + 1}.xml`] = encoder.encode(slide);
  }
  return zipSync(files);
}
