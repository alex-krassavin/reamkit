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
