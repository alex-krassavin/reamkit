// E-PDF EP5 — the pdfReader adapter. Implements the DocumentReader contract so
// `Ream.parse(pdfBytes)` sniffs %PDF-, parses the objects (EP1), extracts text
// (EP2) and reconstructs a FlowDoc — via the tagged structure tree when present
// (EP3) or the layout heuristic otherwise (EP4). Raster images are lifted back
// out and placed in reading order (EP6); true vector graphics are not.

import { PdfFile } from './document';
import { collectPageVectors } from './vector';
import { extractPageText } from './text';
import { reconstructByLayout } from './layout';
import { reconstructTaggedPdf } from './tagged';
import type { PdfVector } from './vector';
import type { StreamFilters } from './document';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';

import { FEATURES } from '@/core/ir';
import { WrongPasswordError } from '@/core/crypto/password-error';

// Sniff the %PDF- header, tolerating a few junk bytes before it (some producers
// prepend a BOM or whitespace) — §7.5.2.
function sniffPdf(bytes: Uint8Array): boolean {
  const limit = Math.min(bytes.length - 5, 1024);
  for (let i = 0; i <= limit; i++) {
    if (
      bytes[i] === 0x25 && // %
      bytes[i + 1] === 0x50 && // P
      bytes[i + 2] === 0x44 && // D
      bytes[i + 3] === 0x46 && // F
      bytes[i + 4] === 0x2d // -
    ) {
      return true;
    }
  }
  return endsLikePdf(bytes);
}

/**
 * §7.5.5 — a file with no header at all, recognised by how it ENDS.
 *
 * bug1606566.pdf begins with the binary comment that normally FOLLOWS the
 * header and has no `%PDF-` anywhere: the producer wrote the second line and
 * not the first. Every reader takes it — poppler says "May not be a PDF file
 * (continuing anyway)" and reads its one line of text — and we refused the file
 * outright, which is the worst answer of the three.
 *
 * The two tokens that close every PDF and close nothing else: the cross-
 * reference offset and the end-of-file marker, in that order, at the end.
 */
function endsLikePdf(bytes: Uint8Array): boolean {
  const tail = bytes.subarray(Math.max(0, bytes.length - 2048));
  const text = new TextDecoder('latin1').decode(tail);
  const eof = text.lastIndexOf('%%EOF');
  return eof >= 0 && text.lastIndexOf('startxref') < eof;
}

/**
 * Parse PDF bytes into a {@link FlowDoc} (E-PDF EP5): reconstruct via the tagged
 * structure tree when present (EP3), else the layout heuristic (EP4), lifting
 * raster images back into reading order (EP6). Records losses for encryption that
 * could not be opened, heuristic (untagged) reconstruction, and the vector
 * regions that are not reconstructed (clipping paths, bare `sh` shadings).
 *
 * @param bytes    The complete PDF file bytes.
 * @param password The user password for an encrypted source; the empty string
 *                 opens permissions-only encryption.
 * @param filters  Decoders for `/Filter` names this reader does not implement
 *                 (§7.4); see {@link StreamFilters}.
 * @returns The reconstructed FlowDoc and its accumulated {@link Loss} report.
 * @throws WrongPasswordError when the file is encrypted and the password opens
 *         it neither as its user nor as its owner — none given included.
 *
 * Which of the two readings a file gets is the FILE's to decide and no
 * caller's: a page that is mostly marks is reproduced where it stands, one
 * that is mostly lines is re-set as a document, and the reader records which
 * it chose. There is no override — see the note on {@link readingOf}. A file
 * kept as a page also carries its `flowing` reading, for the one medium that
 * can keep no page: see {@link ReadResult.flowing}.
 */
export function readPdf(
  bytes: Uint8Array,
  password = '',
  filters: StreamFilters = {},
): ReadResult<FlowDoc> {
  const file = PdfFile.parse(bytes, password, filters);
  // §7.6.3.4 — every revision of the handler stores a verifier, so a password
  // that opens the file as neither its user nor its owner is KNOWN to be wrong,
  // and is refused as an Office package's is. It used to come back as a
  // document with no text and a line in the losses, which a caller that reads
  // the text and not the losses takes for an empty file.
  if (file.passwordRefused) {
    throw new WrongPasswordError(
      password === ''
        ? 'This PDF needs a password to open — pass it as `password`'
        : 'The password opens this PDF neither as its user nor as its owner',
    );
  }
  // Which reading the file asks for. The FILE decides, always: a caller cannot
  // know whether the bytes it was handed are a paper or a form, and asking it
  // to choose only moved the guess outward — and left the choice the library
  // actually makes untested, because everything that measured it pinned the
  // other one.
  const { reading, forText } = readingOf(file);
  const read = readAs(file, reading);
  if (!forText) return read;
  // Letters the file names no character for are DRAWN, and a page of them is
  // kept where they stand: that is the only place a medium that shows them —
  // as pictures of words — can put them. A page ruled into a table is kept
  // where its rules stand, and its cells with them, which is how a paged medium
  // shows the table exactly. Markdown shows neither the letters nor the rules.
  // It drops them, and from a page it had the readable lines left over a line
  // to a paragraph: TAMReview.pdf, whose body is set in such letters, came back
  // with its abstract cut at every line, and a ruled estimate with every cell a
  // paragraph of its own and "1 250 000" as "1 250" and "000". So that medium
  // reads the file again as the document it is without them.
  //
  // Only then. A form, a grid or a drawing is a page for its ARTWORK, which a
  // flowing reading reorders: read as a document, a form's fields came back
  // before their labels and a grid's repeated captions were taken for running
  // heads, which markdown drops too. Parsed afresh rather than held: the
  // file's caches would otherwise outlive every conversion that never asks.
  return { ...read, flowing: () => readAs(PdfFile.parse(bytes, password, filters), 'flow', true) };
}

/**
 * The file read the one way: the losses every reading reports, the reading's
 * own, and its reconstruction.
 *
 * @param file     The parsed file.
 * @param reading  Whether the words re-set as a document or stand where drawn.
 * @param reflowed Whether this is the flowing reading of a file the reader
 *                 keeps as a page, for a medium that places nothing.
 * @returns The reconstructed FlowDoc and its losses.
 */
function readAs(
  file: PdfFile,
  reading: 'flow' | 'positional',
  reflowed = false,
): ReadResult<FlowDoc> {
  const losses: Array<Loss> = [];

  if (file.encryptionUnsupported) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail:
        'encrypted PDF — its security handler is not supported, or its /Encrypt dictionary cannot be read, so its content cannot be decrypted',
    });
  }

  // §7.4 — a stream filter nothing here can undo leaves that stream unread.
  // Reported before anything else: when it is the cross-reference stream, every
  // page of the file is missing and no other loss explains why.
  for (const name of file.unknownFilters) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail: `PDF stream filter /${name} is not supported; streams using it are unreadable`,
    });
  }

  losses.push({
    severity: 'degraded',
    feature: FEATURES.text,
    detail: reflowed
      ? 'PDF read as a DOCUMENT (flowing) for a medium that keeps only the text: the file is kept as a PAGE elsewhere for its drawn letters or its rules, which only a page can show where they stand'
      : reading === 'positional'
        ? 'PDF read as a PAGE (placed): its artwork outweighs its prose, so every line stands where its glyphs stand'
        : 'PDF read as a DOCUMENT (flowing): its prose outweighs its artwork, so the words re-set and the artwork takes its turn in reading order',
  });
  // A placed reconstruction never consults the structure tree: the tree names a
  // reading order, and a reading order is the one thing a placed page does not
  // have. The words go where the glyphs are.
  const tagged = reading === 'positional' ? undefined : reconstructTaggedPdf(file);
  const reconstruction = tagged ?? reconstructByLayout(file, reading);
  if (!tagged && reading !== 'positional') {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.text,
      detail:
        'untagged PDF — text and headings reconstructed heuristically from glyph positions; structure is approximate',
    });
  }
  // Per-image losses from EP6 (undecodable colour spaces, dropped alpha, …).
  losses.push(...reconstruction.losses);
  return { doc: reconstruction.doc, losses };
}

/**
 * Which reading a PDF asks for: a DOCUMENT to re-set, or a PAGE to reproduce.
 *
 * The two cannot be mixed. A flowing reading moves the words, and words that
 * move cannot agree with rules that do not — anchored artwork over reflowed
 * text puts every label on the wrong box. A placed reading keeps both where
 * they were drawn, and reflows nothing.
 *
 * The file says which it is by what is on it. A form or a drawing is mostly
 * MARKS — 160F-2019.pdf sets 28 numbered rows in 355 ruled boxes — and a paper
 * is mostly LINES, with a rule or two between them. So the marks are counted
 * against the lines, on the MEDIAN page rather than the worst, since one plan
 * folded into a report does not make the report a plan.
 *
 * A threshold is a guess, and this one is stated rather than hidden: the reader
 * records which reading it took and why.
 *
 * A glyph the file names no character for is drawn, and counted as a mark like
 * any other path — it IS one, for every medium that has to place it; so is a
 * rule. Whether the page was a page for those alone is said as well: see
 * `forText`.
 */
function readingOf(file: PdfFile): {
  readonly reading: 'flow' | 'positional';
  /**
   * The file is a page only for marks a medium that keeps nothing but the text
   * drops anyway — its drawn letters and its rules: without them it would flow.
   */
  readonly forText: boolean;
} {
  /**
   * Below this many marks a page is prose with decoration, whatever the ratio.
   *
   * It stood at twenty, and calgray.pdf is a five-by-four grid of grey swatches
   * with a label in each: twenty boxes, of which the one painted white is not a
   * mark anybody can see. Nineteen — one short — and all three pages of it were
   * read as prose, the labels of each row run together into a line and the
   * sheet spilling onto a second page. Nineteen boxes in a grid are not a page
   * of prose with a rule under its heading.
   */
  const ENOUGH_MARKS = 12;
  /** Twice as many marks as lines is a page that is drawn rather than written. */
  const DRAWN = 2;
  /** Below this many runs, an angle is a stamp or a watermark and not the page. */
  const ENOUGH_TURNED = 8;
  /**
   * How many rules a page may have for each mark of other artwork and still be
   * a ruled document: an estimate shades its heading row and its total, and a
   * dozen shaded cells under eighty rules are a table's, not a drawing's.
   */
  const RULES_PER_MARK = 4;
  const ratios: Array<number> = [];
  /** The same, the drawn letters left out. */
  const artwork: Array<number> = [];
  const ratioOf = (marks: number, lines: number): number =>
    marks < ENOUGH_MARKS ? 0 : lines > 0 ? marks / lines : DRAWN;
  for (const page of file.pages()) {
    let marks = 0;
    let letters = 0;
    let rules = 0;
    let lines = 0;
    let turned = 0;
    let runs = 0;
    try {
      const vectors = collectPageVectors(file, page, []).vectors;
      marks = vectors.length;
      letters = vectors.filter((v) => v.glyph === true).length;
      rules = vectors.filter(isRule).length;
      // Baselines, not runs: a line broken into twenty runs is still one line,
      // and counting runs would make ordinary justified prose look drawn.
      const ys = new Set<number>();
      for (const run of extractPageText(file, page)) {
        ys.add(Math.round(run.y));
        runs++;
        if (run.angleDeg !== undefined) turned++;
      }
      lines = ys.size;
    } catch {
      continue;
    }
    // §9.4.2 — a page whose words are set at an ANGLE is a page being drawn,
    // whatever else is on it: the placement IS the content, and re-set flat the
    // words come back in an order the page never had. bug946506.pdf runs every
    // line of its lorem ipsum down the sheet at twenty degrees, and read as
    // prose its columns interleaved — "adipiscinnon luctus eleipsum dolor sit".
    if (runs >= ENOUGH_TURNED && turned > runs * 0.5) {
      ratios.push(DRAWN);
      artwork.push(DRAWN);
      continue;
    }
    ratios.push(ratioOf(marks, lines));
    // …and for a medium that keeps only the TEXT, which marks of the page are
    // still a reason to keep it a page. Not its drawn letters (see `forText`).
    // Not its RULES either: a rule is the text's punctuation — the lines round a
    // table's cells, under a heading, over a total — and such a medium drops it.
    // An estimate set as one ruled table is 84 strokes to 9 lines, where
    // 160F-2019.pdf's form is 223 marks to 93; kept a page, every cell of it
    // came back a line of its own and "1 250 000" as "1 250" and "000". A
    // form's fields and a grid's swatches are boxes, and stay; a page with no
    // text is a drawing, rules or not.
    const ruled =
      lines > 0 && marks - letters - rules < Math.max(ENOUGH_MARKS, rules / RULES_PER_MARK);
    artwork.push(ruled ? 0 : ratioOf(marks - letters, lines));
  }
  const drawn = (rs: ReadonlyArray<number>): boolean => {
    const sorted = [...rs].sort((a, b) => a - b);
    return (sorted[Math.floor(sorted.length / 2)] ?? 0) >= DRAWN;
  };
  if (!drawn(ratios)) return { reading: 'flow', forText: false };
  return { reading: 'positional', forText: !drawn(artwork) };
}

/** Thinner than this a straight mark is a line, whatever it was drawn as. */
const RULE_THICKNESS = 2.5;
/** …and it is a rule rather than a dot or a tick from this long. */
const RULE_LENGTH = 4;

/**
 * Whether a painted path is a RULE: straight, along the page, a line thin
 * enough to be one — stroked as a line or filled as a hairline box, which is
 * how many producers draw a table's grid.
 */
function isRule(v: PdfVector): boolean {
  if (v.glyph === true || v.gradient !== undefined) return false;
  if (v.segs.some((seg) => seg.op === 'cubic')) return false;
  const w = v.maxX - v.minX;
  const h = v.maxY - v.minY;
  return Math.min(w, h) <= RULE_THICKNESS && Math.max(w, h) >= RULE_LENGTH;
}

/**
 * The `pdfReader` adapter: a {@link DocumentReader} that sniffs the `%PDF-`
 * header and parses the bytes into a {@link FlowDoc} (E-PDF EP5).
 */
export const pdfReader: DocumentReader<FlowDoc> = {
  id: 'pdf',
  produces: 'flow',
  supports: new Set([FEATURES.text, FEATURES.tables, FEATURES.lists, FEATURES.images]),
  sniff: sniffPdf,
  read: (bytes, opts) =>
    readPdf(
      bytes,
      typeof opts?.password === 'string' ? opts.password : '',
      isFilters(opts?.filters) ? opts.filters : {},
    ),
};

/** A caller's `filters` option, when it is the shape the reader can use. */
function isFilters(value: unknown): value is StreamFilters {
  return typeof value === 'object' && value !== null;
}
