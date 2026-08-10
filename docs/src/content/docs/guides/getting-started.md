---
title: Getting started
description: Install Ream and convert your first document to PDF.
---

Ream reads Word, Excel, PowerPoint and PDF — the modern `.docx` / `.xlsx` / `.pptx` /
`.pdf` and the legacy binary `.doc` / `.xls` / `.ppt` — and converts any of them to
PDF, SVG, HTML, Markdown, DOCX or XLSX, implemented from the ECMA-376 and ISO 32000
specifications. It works on `Uint8Array` in and `Uint8Array` out, so the same code
runs in the browser, Node.js, serverless and edge runtimes.

## Install

The package is published as `reamkit`:

```sh
npm install reamkit
```

Runtime dependencies are minimal: `fflate` (ZIP/Deflate) and `fast-xml-parser`.

## Convert a document

Parse once into the format-neutral interlayer, then convert to any target.
The format (docx/xlsx/pptx/pdf) is sniffed from the bytes. No fonts to wire up — an
open metric-compatible substitute (Arimo for sans, Tinos for serif, Cousine for
monospace, plus Carlito/Caladea for Calibri/Cambria — the same families LibreOffice
substitutes) is fetched automatically based on the document's referenced fonts.
Those five are Latin, so a document that holds **Japanese, Korean, Chinese,
Arabic, Hebrew, Thai or geometric symbols** also gets a Noto face for that
writing system — one weight, fetched only for the scripts the document actually
uses, and a font the document **embeds** is used ahead of any of them:

```ts
import { Ream } from 'reamkit';

// e.g. from an <input type="file"> or a fetch() — anything that yields bytes.
const bytes = new Uint8Array(await file.arrayBuffer());

const doc = Ream.parse(bytes);          // docx, xlsx, pptx or pdf — sniffed
const pdf = await doc.convert('pdf');   // async — fetches a font if needed
const svg = await doc.convert('svg');   // same parse, different target
const html = await doc.convert('html');  // flowed HTML — needs no fonts at all
const md   = await doc.convert('md');    // GitHub-Flavored Markdown — same, narrower
const docx = await doc.convert('docx');  // WordprocessingML back out
const xlsx = await doc.convert('xlsx');  // SpreadsheetML back out (xlsx source)

// Hand the bytes to the browser: preview, download, upload, …
const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
window.open(url);
```

## Open a password-protected document

An encrypted `.docx`/`.xlsx` is not a zip: ECMA-376 §2.3 puts the whole package
inside an OLE container, so it has to be decrypted before anything can read it.
Pass the password to `parse` and that happens for you — both Office schemes
(the 2007 standard one and the agile one 2010 and later write), and the same
option carries a PDF's user password:

```ts
const doc = Ream.parse(bytes, { password: 'letmein' });
const pdf = await doc.convert('pdf');
```

Every scheme stores a verifier, so a wrong password is refused rather than
producing rubbish — as `WrongPasswordError`, which you can catch by type and
so tell apart from a corrupt file:

```ts
import { Ream, WrongPasswordError } from 'reamkit';

try {
  Ream.parse(bytes, { password });
} catch (e) {
  if (e instanceof WrongPasswordError) promptAgain();
  else throw e;
}
```

A document that needs a password and gets none throws as well, saying so. To
ask before that happens — to put up a password prompt instead of an error —
`isEncryptedPackage` answers from the bytes alone:

```ts
import { isEncryptedPackage } from 'reamkit';

const password = isEncryptedPackage(bytes) ? await promptForPassword() : undefined;
const doc = Ream.parse(bytes, { password });
```

## Bring your own fonts (no network)

To embed specific fonts — or to avoid the network entirely — pass the font
bytes in. `convert` then performs zero I/O:

```ts
const fonts = {
  regular: new Uint8Array(await fetch('/fonts/MyFont-Regular.ttf').then((r) => r.arrayBuffer())),
  bold: new Uint8Array(await fetch('/fonts/MyFont-Bold.ttf').then((r) => r.arrayBuffer())),
};

const pdf = await Ream.parse(bytes).convert('pdf', { fonts });
```

## The document object

```ts
const doc = Ream.parse(bytes);
doc.format;    // 'docx' | 'xlsx' | 'pptx' | 'pdf'
doc.flow;      // the parsed interlayer tree (paragraphs, tables, images, …)
doc.losses;    // anything dropped/degraded while reading

const { bytes: out, losses } = await doc.convertWithReport('pdf', { fonts });
await doc.convert('pdf', { fonts, strict: true }); // throw on the first loss
```

## Next steps

- [Examples](/guides/examples/) — PDF/A, signatures, font providers,
  SVG preview, recipes.
- [Concepts](/guides/concepts/) — the pipeline and the interlayer.
