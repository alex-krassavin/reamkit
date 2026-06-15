// Build a minimal valid PresentationML pptx in memory for tests (E-PPTX). Emits
// what the reader needs: [Content_Types].xml, _rels/.rels, ppt/presentation.xml
// (slide size + slide order) and one ppt/slides/slideN.xml per slide, with the
// matching part relationships.
//
// With `layoutMaster` set it additionally emits one slideLayout + slideMaster,
// wires every slide → layout → master, and lets a slide carry placeholder
// shapes whose geometry/text styles resolve through that cascade (PX2).

import { zipSync } from 'fflate';

const encoder = new TextEncoder();

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const LAYOUT_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const MASTER_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const THEME_CT = 'application/vnd.openxmlformats-officedocument.theme+xml';

export interface BuildPptxLayoutMaster {
  /** Inner XML of the layout's `<p:spTree>` (placeholder shapes with geometry). */
  readonly layoutSpTree?: string;
  /** Inner XML of the master's `<p:spTree>`. */
  readonly masterSpTree?: string;
  /** The master's `<p:txStyles>…</p:txStyles>` block (per-level defaults). */
  readonly txStyles?: string;
  /** Inner XML of the theme's `<a:clrScheme>` (slot colours); wires master → theme. */
  readonly theme?: string;
}

export interface BuildPptxOptions {
  /** Slide size in EMU (default a 16:9 deck, 13.333" × 7.5"). */
  readonly cx?: number;
  readonly cy?: number;
  /** Emit a slideLayout + slideMaster wired to every slide (PX2 cascade). */
  readonly layoutMaster?: BuildPptxLayoutMaster;
  /** Extra package parts (path → bytes), e.g. media images (PX3). */
  readonly media?: Record<string, Uint8Array>;
  /** Per-slide extra `<Relationship/>` XML appended to that slide's .rels. */
  readonly slideRels?: ReadonlyArray<string>;
}

const NS = `xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"`;
const EXT_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
};

// `slides[i]` is the inner XML of that slide's `<p:spTree>` (default empty).
export function buildPptx(
  slides: ReadonlyArray<string>,
  options: BuildPptxOptions = {},
): Uint8Array {
  const cx = options.cx ?? 12192000;
  const cy = options.cy ?? 6858000;
  const n = Math.max(1, slides.length);
  const lm = options.layoutMaster;

  // Content-type Defaults for any media extensions present (png/jpg/…).
  const mediaExts = new Set(
    Object.keys(options.media ?? {})
      .map((p) => p.split('.').pop()?.toLowerCase())
      .filter((e): e is string => e !== undefined && e !== 'rels' && e !== 'xml'),
  );
  const mediaDefaults = [...mediaExts]
    .map(
      (e) =>
        `<Default Extension="${e}" ContentType="${EXT_MIME[e] ?? 'application/octet-stream'}"/>`,
    )
    .join('');

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Types xmlns="${PKG_REL_NS.replace('relationships', 'content-types')}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    mediaDefaults +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    Array.from(
      { length: n },
      (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${SLIDE_CT}"/>`,
    ).join('') +
    (lm
      ? `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${LAYOUT_CT}"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${MASTER_CT}"/>`
      : '') +
    (lm?.theme ? `<Override PartName="/ppt/theme/theme1.xml" ContentType="${THEME_CT}"/>` : '') +
    `</Types>`;

  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="ppt/presentation.xml"/>` +
    `</Relationships>`;

  const presentation =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<p:presentation ${NS}>` +
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
      `<p:sld ${NS}>` +
      `<p:cSld><p:spTree>${slides[i] ?? ''}</p:spTree></p:cSld>` +
      `</p:sld>`;
    files[`ppt/slides/slide${i + 1}.xml`] = encoder.encode(slide);
    // Slide .rels: the layout link (when a layout/master is present) plus any
    // caller-supplied extra relationships (e.g. an image rel into media).
    const rels: Array<string> = [];
    if (lm) {
      rels.push(
        `<Relationship Id="rIdLayout" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
      );
    }
    const extra = options.slideRels?.[i];
    if (extra) rels.push(extra);
    if (rels.length > 0) {
      files[`ppt/slides/_rels/slide${i + 1}.xml.rels`] = encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<Relationships xmlns="${PKG_REL_NS}">${rels.join('')}</Relationships>`,
      );
    }
  }

  if (lm) {
    files['ppt/slideLayouts/slideLayout1.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<p:sldLayout ${NS}><p:cSld><p:spTree>${lm.layoutSpTree ?? ''}</p:spTree></p:cSld></p:sldLayout>`,
    );
    files['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${R_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    );
    files['ppt/slideMasters/slideMaster1.xml'] = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<p:sldMaster ${NS}><p:cSld><p:spTree>${lm.masterSpTree ?? ''}</p:spTree></p:cSld>` +
        `${lm.txStyles ?? ''}</p:sldMaster>`,
    );
    if (lm.theme) {
      files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<Relationships xmlns="${PKG_REL_NS}">` +
          `<Relationship Id="rId1" Type="${R_NS}/theme" Target="../theme/theme1.xml"/>` +
          `</Relationships>`,
      );
      files['ppt/theme/theme1.xml'] = encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<a:theme xmlns:a="${A_NS}" name="deck"><a:themeElements>` +
          `<a:clrScheme name="deck">${lm.theme}</a:clrScheme>` +
          `</a:themeElements></a:theme>`,
      );
    }
  }
  for (const [path, bytes] of Object.entries(options.media ?? {})) files[path] = bytes;
  return zipSync(files);
}
