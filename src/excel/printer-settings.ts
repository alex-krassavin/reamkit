// The `printerSettings#.bin` part a `<pageSetup r:id>` points at: a Windows
// DEVMODE structure (MS-RPRN §2.2.2.1) as saved by the print dialog.
//
// It matters because `<pageSetup>` may name no `paperSize` at all and still not
// mean "the default": Excel records the choice here instead, and LibreOffice
// reads it. simple-monthly-budget.xlsx and 45540_classic_Header.xlsx both do
// that, and both print on Letter while we assumed A4 — the single most visible
// difference on every page of either document.

/** The paper size and orientation a DEVMODE records, when it records them. */
export interface PrinterPageSetup {
  /** `dmPaperSize` — the same §18.3.1.63 enumeration `<pageSetup paperSize>` uses. */
  readonly paperSize?: number;
  readonly orientation?: 'portrait' | 'landscape';
}

// DEVMODE, Unicode form: dmDeviceName[32] WCHARs, then the fixed header. The
// two fields we want sit at 76 and 78, and dmFields says whether each was set.
const DEVICE_NAME_BYTES = 64;
const OFFSET_SIZE = DEVICE_NAME_BYTES + 4;
const OFFSET_ORIENTATION = DEVICE_NAME_BYTES + 12;
const OFFSET_PAPER_SIZE = DEVICE_NAME_BYTES + 14;
const DM_ORIENTATION = 0x0000_0001;
const DM_PAPERSIZE = 0x0000_0002;

/**
 * Read the paper size and orientation out of a `printerSettings#.bin` part.
 *
 * Returns an empty record for anything that does not look like a DEVMODE — the
 * part is opaque binary from an arbitrary printer driver, so nothing here may
 * assume more than the fixed header it is required to start with.
 *
 * @param data The raw part bytes.
 */
export function parsePrinterSettings(data: Uint8Array): PrinterPageSetup {
  if (data.byteLength < OFFSET_PAPER_SIZE + 2) return {};
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const dmSize = view.getUint16(OFFSET_SIZE, true);
  // dmSize covers the fixed header; a value that cannot reach the fields we
  // read means this is not the structure we think it is.
  if (dmSize < OFFSET_PAPER_SIZE + 2 || dmSize > data.byteLength) return {};
  const fields = view.getUint32(DEVICE_NAME_BYTES + 8, true);
  const out: { -readonly [K in keyof PrinterPageSetup]: PrinterPageSetup[K] } = {};
  if (fields & DM_PAPERSIZE) {
    const paper = view.getInt16(OFFSET_PAPER_SIZE, true);
    // Negative values are the driver's own custom forms, which name no size.
    if (paper > 0) out.paperSize = paper;
  }
  if (fields & DM_ORIENTATION) {
    const orientation = view.getInt16(OFFSET_ORIENTATION, true);
    if (orientation === 1) out.orientation = 'portrait';
    else if (orientation === 2) out.orientation = 'landscape';
  }
  return out;
}
