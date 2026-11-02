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
| `Arimo-Regular.ttf` | `02219a6ff8456d98` | `Arimo_400Regular.ttf` from `@expo-google-fonts/arimo` on jsDelivr — the face the library downloads for a document that names no font. `tests/ligatures.test.ts` turns on its coverage: no U+FB00, U+FB01/U+FB02 with their private-use twins U+F001/U+F002, and U+210E. |

## Noto Sans Devanagari 2.007, Bengali 3.011, Kannada 2.006, Newa 2.007 — OFL-1.1

Copyright 2022–2025 The Noto Project Authors (<https://github.com/notofonts>).
The licence, with each face's copyright line, is reproduced in `Noto-OFL.txt`;
none of the four reserves its font name.

Each is the `_400Regular.ttf` of its `@expo-google-fonts/noto-sans-*` package
on jsDelivr — for the first three, the face the library downloads for that
script — cut down with
`hb-subset` (HarfBuzz 15.0.0) to the characters `tests/opentype-lookups.test.ts`
sets, every layout feature kept: `hb-subset <face> --unicodes=<the characters>
--layout-features='*'`. The glyph ids, clusters and positions that test expects
are what `hb-shape` 15.0.0 gives on these same files.

| File | sha256 (16) | Why it is here |
|---|---|---|
| `NotoSansDevanagari-subset.ttf` | `0817a4cf2744182a` | हिन्दी श्री विद्यालय क्षमा हूँ कुं क़ुँ राष्ट्र द्वार र्कि ज्ञान, ZWJ and ZWNJ: conjuncts by ligature and by chained context, nukta, the presentation stage, marks on their base, pair kerning, join controls. |
| `NotoSansBengali-subset.ttf` | `5ec728f1d69f412d` | সংস্কৃতি হৃদয় চাঁদ ক্তৃ: a mark on a mark, and positioning by context. |
| `NotoSansKannada-subset.ttf` | `38e777f33682f9dd` | ಙೈ and the two parts of its vowel sign, U+0CC6 U+0CD6: a pair whose second glyph starts the next pair. |
| `NotoSansNewa-subset.ttf` | `53d77b12fd169a0c` | Four Newa syllables: cursive attachment. Of the faces the library downloads only Kannada's has a `curs` feature, and none of its joins moves a glyph. |
