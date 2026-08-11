// §7.10 — the four kinds of PDF function, and the tint transform that is the
// reason they can be called at all. A `/Separation` or `/DeviceN` fill whose
// transform is not run has only its tint to go on, and a tint of 1 read as a
// grey level is BLACK: devicen.pdf's three triangles are green, blue and red,
// and all three came back black.

import { describe, expect, it } from 'vitest';

import { PdfFile } from '@/pdf-reader/document';
import { PdfRef } from '@/pdf/objects';
import { readFunction } from '@/pdf-reader/function';
import { Ream } from '@/core/converter/ream';

/** A one-page file whose objects are given as bodies, `content` its stream. */
function pdfWith(
  objects: ReadonlyArray<string>,
  streams: ReadonlyMap<number, Uint8Array>,
): Uint8Array {
  const bytes: Array<number> = [];
  const push = (s: string): void => {
    for (const ch of s) bytes.push(ch.charCodeAt(0) & 0xff);
  };
  push('%PDF-1.7\n');
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(bytes.length);
    push(`${String(i + 1)} 0 obj\n${body}\n`);
    const data = streams.get(i + 1);
    if (data) {
      push('stream\n');
      for (const b of data) bytes.push(b);
      push('\nendstream\n');
    }
    push('endobj\n');
  });
  const xref = bytes.length;
  push(`xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`);
  for (const off of offsets) push(`${String(off).padStart(10, '0')} 00000 n \n`);
  push(
    `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`,
  );
  return Uint8Array.from(bytes);
}

/** A file holding one function as object 4, and that function, read. */
function loneFunction(body: string, data?: Uint8Array): ReturnType<typeof readFunction> {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>',
    body,
  ];
  const file = PdfFile.parse(pdfWith(objects, new Map(data ? [[4, data]] : [])));
  return readFunction(file, new PdfRef(4, 0));
}

/** …reached by object number, which is what a colour space states. */
function fnOf(body: string, data?: Uint8Array): NonNullable<ReturnType<typeof readFunction>> {
  const got = loneFunction(body, data);
  expect(got).toBeDefined();
  return got!;
}

describe('§7.10.3 — an exponential function', () => {
  it('interpolates between C0 and C1', () => {
    const fn = fnOf('<< /FunctionType 2 /Domain [0 1] /C0 [0 0.2 0.4] /C1 [1 0.6 0.4] /N 1 >>');
    expect(fn([0])).toEqual([0, 0.2, 0.4]);
    expect(fn([1])).toEqual([1, 0.6, 0.4]);
    const mid = fn([0.5]);
    expect(mid[0]).toBeCloseTo(0.5, 6);
    expect(mid[1]).toBeCloseTo(0.4, 6);
  });

  it('bends it by the exponent', () => {
    const fn = fnOf('<< /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 2 >>');
    expect(fn([0.5])[0]).toBeCloseTo(0.25, 6);
  });

  it('clips an input to the domain rather than running off it', () => {
    const fn = fnOf('<< /FunctionType 2 /Domain [0 0.5] /C0 [0] /C1 [1] /N 1 >>');
    expect(fn([9])[0]).toBeCloseTo(0.5, 6);
  });
});

describe('§7.10.4 — a stitching function', () => {
  const body =
    '<< /FunctionType 3 /Domain [0 1] /Bounds [0.5] /Encode [0 1 0 1] /Functions [' +
    '<< /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 1 >> ' +
    '<< /FunctionType 2 /Domain [0 1] /C0 [1] /C1 [0] /N 1 >>] >>';

  it('picks the subfunction the input falls in, and re-encodes onto it', () => {
    const fn = fnOf(body);
    // First half runs 0 → 1 across [0, 0.5); second runs 1 → 0 across [0.5, 1].
    expect(fn([0])[0]).toBeCloseTo(0, 6);
    expect(fn([0.25])[0]).toBeCloseTo(0.5, 6);
    expect(fn([0.5])[0]).toBeCloseTo(1, 6);
    expect(fn([0.75])[0]).toBeCloseTo(0.5, 6);
    expect(fn([1])[0]).toBeCloseTo(0, 6);
  });
});

describe('§7.10.2 — a sampled function', () => {
  it('reads the table and interpolates between its samples', () => {
    // Three 8-bit samples, one output: 0, 128, 255.
    const fn = fnOf(
      '<< /FunctionType 0 /Domain [0 1] /Range [0 1] /Size [3] /BitsPerSample 8 /Length 3 >>',
      Uint8Array.from([0, 128, 255]),
    );
    expect(fn([0])[0]).toBeCloseTo(0, 6);
    expect(fn([0.5])[0]).toBeCloseTo(128 / 255, 6);
    expect(fn([1])[0]).toBeCloseTo(1, 6);
    // A quarter of the way is halfway between the first two samples.
    expect(fn([0.25])[0]).toBeCloseTo(64 / 255, 6);
  });

  it('runs the first input fastest, over more than one dimension', () => {
    // A 2×2 table of one output: [ (0,0)=0  (1,0)=1  (0,1)=2  (1,1)=3 ] / 3.
    const fn = fnOf(
      '<< /FunctionType 0 /Domain [0 1 0 1] /Range [0 1] /Size [2 2] /BitsPerSample 8 /Length 4 >>',
      Uint8Array.from([0, 85, 170, 255]),
    );
    expect(fn([0, 0])[0]).toBeCloseTo(0, 6);
    expect(fn([1, 0])[0]).toBeCloseTo(1 / 3, 2);
    expect(fn([0, 1])[0]).toBeCloseTo(2 / 3, 2);
    expect(fn([1, 1])[0]).toBeCloseTo(1, 6);
    // The middle of the square is the mean of all four corners.
    expect(fn([0.5, 0.5])[0]).toBeCloseTo(0.5, 2);
  });

  it('carries the samples onto /Decode where the file states one', () => {
    const fn = fnOf(
      '<< /FunctionType 0 /Domain [0 1] /Range [0 1] /Decode [1 0] /Size [2] ' +
        '/BitsPerSample 8 /Length 2 >>',
      Uint8Array.from([0, 255]),
    );
    expect(fn([0])[0]).toBeCloseTo(1, 6);
    expect(fn([1])[0]).toBeCloseTo(0, 6);
  });
});

describe('§7.10.5 — a PostScript calculator', () => {
  const ps = (program: string, range = '[0 1]'): NonNullable<ReturnType<typeof readFunction>> =>
    fnOf(
      `<< /FunctionType 4 /Domain [0 1] /Range ${range} /Length ${String(program.length)} >>`,
      Uint8Array.from([...program].map((c) => c.charCodeAt(0))),
    );

  it('runs arithmetic on the stack', () => {
    expect(ps('{ 2 mul 0.1 add }')([0.4])[0]).toBeCloseTo(0.9, 6);
  });

  it('takes the LAST n values as the outputs', () => {
    // Three on the stack, three wanted: the input three times over.
    const fn = ps('{ dup dup }', '[0 1 0 1 0 1]');
    expect(fn([0.25])).toEqual([0.25, 0.25, 0.25]);
  });

  it('branches on ifelse', () => {
    const fn = ps('{ 0.5 gt { 1 } { 0 } ifelse }');
    expect(fn([0.9])[0]).toBe(1);
    expect(fn([0.1])[0]).toBe(0);
  });

  it('branches on a bare if, which leaves the stack alone when false', () => {
    const fn = ps('{ dup 0.5 gt { pop 1 } if }');
    expect(fn([0.9])[0]).toBe(1);
    expect(fn([0.2])[0]).toBeCloseTo(0.2, 6);
  });

  it('rolls and indexes the stack', () => {
    // 1 2 3 → roll by one → 3 1 2; the last of the three is 2.
    expect(ps('{ pop 1 2 3 3 1 roll }', '[0 10]')([0])[0]).toBe(2);
    // `index` copies the k-th value down without disturbing it.
    expect(ps('{ pop 7 8 1 index }', '[0 10]')([0])[0]).toBe(7);
  });

  it('clips the outputs to /Range', () => {
    expect(ps('{ 100 mul }')([1])[0]).toBe(1);
  });

  it('gives nothing back for a program it cannot parse', () => {
    expect(
      loneFunction('<< /FunctionType 4 /Domain [0 1] /Range [0 1] /Length 3 >>'),
    ).toBeUndefined();
  });
});

describe('§8.6.6.5 — a DeviceN fill takes the colour its transform gives', () => {
  /** A page filling one square in a DeviceN space at the stated tints. */
  const filled = (tints: string): Uint8Array => {
    // Three colorants named Red/Green/Blue whose transform hands back the
    // components REVERSED — which is the point: the names say nothing, only
    // the transform does. devicen.pdf plays the same trick.
    const ps = '{ 3 1 roll exch }';
    const content = `/Cs cs ${tints} scn 10 10 80 80 re f`;
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R ' +
        '/Resources << /ColorSpace << /Cs 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>`,
      '[ /DeviceN [ /Red /Green /Blue ] /DeviceRGB 6 0 R ]',
      `<< /FunctionType 4 /Domain [0 1 0 1 0 1] /Range [0 1 0 1 0 1] /Length ${String(ps.length)} >>`,
    ];
    const enc = (s: string): Uint8Array => Uint8Array.from([...s].map((c) => c.charCodeAt(0)));
    return pdfWith(
      objects,
      new Map([
        [4, enc(content)],
        [6, enc(ps)],
      ]),
    );
  };

  const fillOf = (tints: string): string | undefined => {
    const shape = Ream.parse(filled(tints)).flow.body.find((el) => el.kind === 'shape');
    return shape?.kind === 'shape' && shape.shape.fill.kind === 'solid'
      ? shape.shape.fill.colorHex
      : undefined;
  };

  it('runs the transform rather than reading the tint as ink', () => {
    // `1 0 0` on colorants called Red/Green/Blue, reversed by the transform,
    // is BLUE. Read as a strength of ink it would be black.
    expect(fillOf('1 0 0')).toBe('0000FF');
    expect(fillOf('0 0 1')).toBe('FF0000');
    expect(fillOf('0 1 0')).toBe('00FF00');
  });

  it('still reads a tint as ink where the transform cannot be run', () => {
    // No /TintTransform at all: full strength is the colorant at its darkest.
    const content = '/Cs cs 1 scn 10 10 80 80 re f';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R ' +
        '/Resources << /ColorSpace << /Cs 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>`,
      '[ /Separation /Spot /DeviceRGB ]',
    ];
    const enc = (s: string): Uint8Array => Uint8Array.from([...s].map((c) => c.charCodeAt(0)));
    const shape = Ream.parse(pdfWith(objects, new Map([[4, enc(content)]]))).flow.body.find(
      (el) => el.kind === 'shape',
    );
    expect(
      shape?.kind === 'shape' && shape.shape.fill.kind === 'solid'
        ? shape.shape.fill.colorHex
        : undefined,
    ).toBe('000000');
  });
});
