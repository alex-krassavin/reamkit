# Corpus validation — findings & progress

Living tracker for divergences surfaced by the real-document harness
(`scripts/corpus/`), which fetches real third-party documents (Apache POI,
LibreOffice regression corpora, Mozilla pdf.js), converts each through Ream, and
diffs the result against a LibreOffice golden render (text similarity +
font-agnostic geometry + pixel mismatch). For a PDF source the reference is the
original file — a read→render roundtrip.

**Run it:**

```sh
npx tsx scripts/corpus/fetch-corpus.ts --source poi- --limit 12   # poi- | lo- | pdfjs | poi-doc …
CORPUS_DIR=corpus/external/poi-doc npx tsx scripts/corpus/run.ts   # in-process — errors show in the report
CORPUS_ISOLATE_OURS=1 CORPUS_DIR=corpus/external/pdfjs npx tsx scripts/corpus/run.ts  # child timeout (fuzzed PDFs)
```

Status: ✅ fixed & committed · 🛠 in progress · ⬜ todo · 🚫 won't-fix (metric/default artifact)

| ID | Format | Finding | Status |
|----|--------|---------|--------|
| F0 | all legacy | CFB v3 garbage high-dword stream size rejected the file | ✅ `705cd44` |
| F1 | pdf | Page size not preserved — every PDF re-renders at A4 | ✅ `916a758` |
| F2a | all legacy | Embedded OLE object's stream shadows the main document's | ✅ `4656d0b` |
| F2b | doc | Word 6.0/95 (nFib < 105) — no CLX, different FIB | 🚫 out of scope |
| F3 | pptx | Hidden slides (show="0") rendered instead of omitted | ✅ `cebed0d` |
| F3b | ppt | Slide content overflows the fixed canvas into extra pages | ⬜ backlog |
| F4 | doc | Section page size / orientation not read (landscape → portrait) | ⬜ backlog |
| N1 | doc/xls | Letter-vs-A4 "dims" mismatch | 🚫 default artifact |

---

## ✅ F0 — CFB v3 garbage high-dword stream size  (commit `705cd44`)

A directory entry's stream size is a 64-bit field, but in a v3 container the
high 4 bytes are reserved and MUST be ignored (MS-CFB §2.6.1); Word leaves
non-zero garbage there. `openCfb` read the full 64-bit value → absurd sizes →
size guard rejected the file. Fix: mask to the low 4 bytes for v3. `53446.doc`:
hard parse failure → 16 KB of real text. Affects all three legacy formats (shared
container). Test: `tests/cfb.test.ts` (garbage-high-dword fixture, both FAT paths).

## ⬜ F1 — PDF page size not preserved (MediaBox → render)

**Symptom.** Every PDF re-renders at the default A4 (827×1170 @100dpi) regardless
of the source's MediaBox: `22060_A1_01_Plans.pdf` is A3 (1170×1654) and our A4
re-render splits its 1 page into 5; Letter / landscape / tiny-custom pages are
all forced to A4 too.

**Root cause.** `src/pdf-reader/document.ts` captures each page's `mediaBox` and
`layout.ts` uses it for layout-internal coordinates, but the FlowDoc the reader
emits never sets a `section`/`sections` page size, so the renderer falls back to
its A4 default.

**Plan.** Mirror the working docx path (`w:pgSz` → section geometry): set the
FlowDoc section width/height from the page MediaBox. Uniform-size PDFs → one
`section`; pages of differing size → multi-`sections`. Thread through both reader
paths (tagged struct-tree + heuristic).

**Validation.** Re-run `pdfjs`: the `dims A vs B` notes collapse (ours == source)
and over-paginated pages (A3 5→1) drop. Add a fixture test: a 2-page PDF with an
A4 page then a landscape page → FlowDoc carries both sizes.

## ✅ F2a — embedded-object stream shadows the main document (commit `4656d0b`)

`53379.doc` (1 char vs 36942) yielded a single zero-width space. Root cause: it
embeds an OLE object, so there are two `WordDocument` / `1Table` streams, and the
CFB reader's flat first-wins lookup returned the **embedded** object's `1Table`
(13346 B). The main FIB's CLX offset (`fcClx+lcbClx=63827`) overran it, so the
piece table came up empty. Fix: resolve streams from the root storage's own
children (walk the directory tree, never descend into a child storage), so the
main document's streams win. `53379.doc` → ~40 KB text / 98% TextSim; `53446.doc`
(also F0) → 16 KB / 90%. Affects all three legacy formats. Test:
`tests/cfb.test.ts` (`storages` fixture).

## 🚫 F2b — Word 6.0/95 binary (nFib < 105)

`57843.doc` is `nFib=101` (Word 6.0): no CLX piece table, `fcClx=lcbClx=0`, and a
different FibBase layout. The reader documents Word 6/95 as out of scope and
degrades gracefully (empty body + `text` loss). Supporting it is a separate
feature with real misread risk and no validatable reference to hand — left as a
principled residual.

## ⬜ F4 — .doc section page size / orientation not read

`53379.doc` renders portrait Letter where LibreOffice shows landscape
(`dims 850x1100 vs 1100x850`): the `.doc` reader doesn't read the section's
page width/height/orientation from the SEP (sprmSXaPage / sprmSYaPage /
sprmSBOrientation), so it falls back to the default. Analogous to the docx
`w:pgSz` path and to F1 for PDF. Backlog (separate from the N1 default-paper
artifact, which is a genuine default difference, not an ignored explicit size).

## ✅ F3 — pptx hidden slides rendered (commit `cebed0d`)

Not overflow, as first suspected — the deck has exactly one page per slide.
PowerPoint and LibreOffice omit hidden slides (`p:sld@show="0"`) from a
printed/exported deck; we rendered them, so the page count ran high.
`2411-Performance_Up.pptx` 48→46, `60810.pptx` 28→25 (exactly the hidden-slide
counts) — both now match the golden. Skip `show="0"` slides + record one
`dropped` loss. Test: `tests/pptx-reader.test.ts` (`hiddenSlides` fixture).

## ⬜ F3b — .ppt slide content overflows into extra pages

`23884…ppt` is 37 pages vs LibreOffice's 30. **Verified it is NOT hidden slides**
(walked the record tree: 30 `RT_Slide` containers, 0 with `SSSlideInfoAtom`
fHidden — and the reader correctly enumerates 30 slides). The 7 extra pages are
**render overflow**: `ppt-reader.buildBody` flows the text of *un-anchored*
shapes (no `rectPt` — e.g. the outline-text fallback) in reading order, and on a
slide whose text exceeds the canvas that flow paginates onto a 2nd page. pptx
never hits this because every shape is positioned (floating, out-of-flow).

No clean fix: floating the un-anchored text at a default rect stops the overflow
but makes *multiple* un-anchored shapes overlap (we have no positions for them);
flowing keeps them readable but paginates. PowerPoint/LibreOffice never split a
slide — they autofit/clip text to placeholders, which is a layout-engine
capability (per-section "no internal pagination" / autofit) we don't have. The
gap is cosmetic (content is all present, ~80% text-sim), so this is deferred
rather than fixed with a regression-prone float. Needs the autofit/clip layout
feature, validated visually — not just on page count.

## 🚫 N1 — Letter vs A4 default paper

Many `.doc`/`.xls` files show `dims 850x1100 vs 827x1170` (Letter vs A4). This is
a default-paper difference, not a bug: our Letter default matches Word's; LO
defaults to A4 by locale. Switching the default would be wrong for US documents.
(Distinct from F1, where the PDF source carries an *explicit* size we ignore.)
