// E-PDF — reverse a stream's /Predictor (ISO 32000-1 §7.4.4.4). Both image
// XObjects (image-decode.ts) and cross-reference / object streams (document.ts)
// FlateDecode their data and then undo a PNG (Predictor ≥ 10) or TIFF
// (Predictor 2) predictor; this is the shared, dependency-free math.

export interface PredictorParams {
  readonly predictor: number;
  readonly colors: number;
  readonly bitsPerComponent: number;
  readonly columns: number;
}

export function reversePredictor(data: Uint8Array, p: PredictorParams): Uint8Array {
  if (p.predictor < 2) return data;
  const bpp = Math.max(1, Math.ceil((p.colors * p.bitsPerComponent) / 8));
  const rowBytes = Math.ceil((p.colors * p.bitsPerComponent * p.columns) / 8);
  if (rowBytes <= 0) return data;

  if (p.predictor === 2) {
    // TIFF horizontal differencing (8-bit components only).
    if (p.bitsPerComponent !== 8) return data;
    const rows = Math.floor(data.length / rowBytes);
    const out = data.slice(0, rows * rowBytes);
    for (let r = 0; r < rows; r++) {
      const off = r * rowBytes;
      for (let i = bpp; i < rowBytes; i++)
        out[off + i] = (out[off + i]! + out[off + i - bpp]!) & 0xff;
    }
    return out;
  }

  // PNG predictors: each row is prefixed with a filter-type byte.
  const stride = rowBytes + 1;
  const rows = Math.floor(data.length / stride);
  const out = new Uint8Array(rows * rowBytes);
  let prev = new Uint8Array(rowBytes);
  for (let r = 0; r < rows; r++) {
    const ft = data[r * stride]!;
    const src = r * stride + 1;
    const dst = r * rowBytes;
    for (let i = 0; i < rowBytes; i++) {
      const x = data[src + i]!;
      const a = i >= bpp ? out[dst + i - bpp]! : 0;
      const b = prev[i]!;
      const c = i >= bpp ? prev[i - bpp]! : 0;
      let v: number;
      switch (ft) {
        case 1:
          v = x + a;
          break;
        case 2:
          v = x + b;
          break;
        case 3:
          v = x + ((a + b) >> 1);
          break;
        case 4:
          v = x + paeth(a, b, c);
          break;
        default:
          v = x;
      }
      out[dst + i] = v & 0xff;
    }
    prev = out.subarray(dst, dst + rowBytes);
  }
  return out;
}

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}
