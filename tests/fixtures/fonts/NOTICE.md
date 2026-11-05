# Test fonts

The faces the suite sets text in, checked in so that no test waits on a CDN.
Each remains under its original licence, as its own `name` table states.

## Roboto 2.001047 — Apache-2.0

Copyright 2015 Google Inc. <https://www.apache.org/licenses/LICENSE-2.0>

| File | sha256 (16) | Why it is here |
|---|---|---|
| `Roboto-Regular.ttf` | `56a45233d29f11b4` | The face most tests pass as `fonts`. |
| `Roboto-Bold.ttf` | `61f89f8db49261c2` | Its bold cut. |
| `Roboto-Italic.ttf` | `fa0b17bb4aaac4a1` | Its italic cut. |
| `Roboto-BoldItalic.ttf` | `40083ed54338397c` | Its bold italic cut. |

## Arimo 1.33 — OFL-1.1

Copyright 2020 The Arimo Project Authors (<https://github.com/googlefonts/arimo>).
The licence is reproduced in `Arimo-OFL.txt`.

| File | sha256 (16) | Why it is here |
|---|---|---|
| `Arimo-Regular.ttf` | `02219a6ff8456d98` | `Arimo_400Regular.ttf` from `@expo-google-fonts/arimo` on jsDelivr — the face the library downloads for a document that names no font. `tests/ligatures.test.ts` turns on its coverage: no U+FB00, U+FB01/U+FB02 with their private-use twins U+F001/U+F002, and U+210E. `tests/turkic-cyrillic.test.ts` on its one glyph for the Latin and the Cyrillic schwa (U+0259/U+04D9, U+018F/U+04D8), its composed Cyrillic letters and its `mark` anchors (the acute over а at −284, −340, as `hb-shape` 15.0.0 gives it). |

## Caladea 1.001 — OFL-1.1

Copyright 2012 The Caladea Project Authors (<https://github.com/huertatipografica/Caladea>).
The licence is reproduced in `Caladea-OFL.txt`; the font reserves no name.

| File | sha256 (16) | Why it is here |
|---|---|---|
| `Caladea-subset.ttf` | `b0f5a40461161047` | `Caladea_400Regular.ttf` from `@expo-google-fonts/caladea` — Cambria's substitute — cut to printable ASCII and the combining acute with `hb-subset` 15.0.0 (`--unicodes=20-7E,301 --layout-features='*'`). `tests/turkic-cyrillic.test.ts` turns on what it lacks — the face has no Cyrillic at all — and on the acute it has. |

## Noto Serif 2.015 — OFL-1.1

Copyright 2022 The Noto Project Authors (<https://github.com/notofonts/latin-greek-cyrillic>).
The licence, with its copyright line, is reproduced in `Noto-OFL.txt`.

| File | sha256 (16) | Why it is here |
|---|---|---|
| `NotoSerif-subset.ttf` | `48fb9db4c1d019c5` | `NotoSerif_400Regular.ttf` from `@expo-google-fonts/noto-serif` — the face Cambria's Cyrillic is borrowed from — cut with `hb-subset` 15.0.0 to the Cyrillic letters of `tests/turkic-cyrillic.test.ts` and the combining acute (`--unicodes=<them>,301 --layout-features='*'`), its `mark` anchors kept. |

## Noto Sans of the scripts of India, and of Newa — OFL-1.1

Devanagari 2.007, Bengali 3.011, Gurmukhi 2.004, Gujarati 2.106, Oriya 2.007,
Tamil 2.004, Telugu 2.005, Kannada 2.006, Malayalam 2.104, Newa 2.007.
Copyright 2022–2025 The Noto Project Authors (<https://github.com/notofonts>).
The licence, with each face's copyright line, is reproduced in `Noto-OFL.txt`;
none of them reserves its font name.

Each is the `_400Regular.ttf` of its `@expo-google-fonts/noto-sans-*` package
on jsDelivr — the face the library downloads for that script, Newa apart — cut
down with `hb-subset` (HarfBuzz 15.0.0) to the characters its test sets, every
layout feature kept: `hb-subset <face> --unicodes=<the characters>
--layout-features='*'`. The glyph ids, clusters and positions those tests expect
are what `hb-shape` 15.0.0 gives on these same files.

For `tests/opentype-lookups.test.ts`, the lookups one at a time:

| File | sha256 (16) | Why it is here |
|---|---|---|
| `NotoSansDevanagari-subset.ttf` | `0817a4cf2744182a` | हिन्दी श्री विद्यालय क्षमा हूँ कुं क़ुँ राष्ट्र द्वार र्कि ज्ञान, ZWJ and ZWNJ: conjuncts by ligature and by chained context, nukta, the presentation stage, marks on their base, pair kerning, join controls. |
| `NotoSansBengali-subset.ttf` | `5ec728f1d69f412d` | সংস্কৃতি হৃদয় চাঁদ ক্তৃ: a mark on a mark, and positioning by context. |
| `NotoSansKannada-subset.ttf` | `38e777f33682f9dd` | ಙೈ and the two parts of its vowel sign, U+0CC6 U+0CD6: a pair whose second glyph starts the next pair. |
| `NotoSansNewa-subset.ttf` | `53d77b12fd169a0c` | Four Newa syllables: cursive attachment. Of the faces the library downloads only Kannada's has a `curs` feature, and none of its joins moves a glyph. |

For `tests/indic-shaper.test.ts`, a few words of each script shaped whole — the
characters of the words in that test and of those whose text layer
`tests/indic-pdf.test.ts` reads, the dotted circle, and the parts of every
two-part vowel sign among them:

| File | sha256 (16) | What it is cut to |
|---|---|---|
| `NotoSansDevanagari-shaping.ttf` | `e17528dacc9eb8e2` | A pre-base i, a reph, ksha, pa with ra under it, the nukta (written apart and as one code point), ZWJ and ZWNJ in a conjunct, a vowel sign on nothing, अ with a vowel sign it may not take, a reph over श्व, an o with an anusvara drawn as the aa and a mark. |
| `NotoSansBengali-shaping.ttf` | `39bc10f1af1ced75` | O and au in two parts, a pre-base i, ya-phala, a reph, a ZWJ before two marks, a reph the face sets after the ii by swapping their glyphs. |
| `NotoSansGurmukhi-shaping.ttf` | `39a1046bfc4d80c7` | Sihari, a subjoined ra, a bindi before ii. |
| `NotoSansGujarati-shaping.ttf` | `27410c182b3c91dc` | A pre-base i, a reph, a half form. |
| `NotoSansOriya-shaping.ttf` | `e0e71f0620f62bbb` | O and ai in two parts, a reph after the main consonant, ksha. |
| `NotoSansTamil-shaping.ttf` | `7db7a3fe8fa51dd7` | O, ai and au around their consonant, ksha, shri, a pair kerned across a ZWJ. |
| `NotoSansTelugu-shaping.ttf` | `c07847a3af772da0` | A reph asked for with ZWJ, ksha, shri, the ai length mark. |
| `NotoSansKannada-shaping.ttf` | `b3aeb9c110441786` | A reph after the post-base forms, o, ai and ii in two parts, two below-base forms, Ra-virama-ZWJ, a nukta inside a ligature. |
| `NotoSansMalayalam-shaping.ttf` | `50843d3cd1059eba` | O and ai around their consonant, a pre-base ra, the dot reph, a chillu. |

For `tests/indic-pdf.test.ts`, a syllable whose signs are coloured apart — क ि ृ
ष ् ण र and the dotted circle:

| File | sha256 (16) | What it is cut to |
|---|---|---|
| `NotoSansDevanagari-colour.ttf` | `283bee12256dccb3` | कृष्ण with the half form of ष्, the short i before its consonant, and the reph that joins it. |
