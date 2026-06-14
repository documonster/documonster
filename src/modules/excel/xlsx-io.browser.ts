/**
 * Browser xlsx IO handle accessor. Kept out of `workbook.browser` so the
 * heavy `XLSX` serializer is not a static dependency of the workbook record
 * module (which would create a workbook ↔ xlsx import cycle).
 */
import type { WorkbookData } from "@excel/workbook-core";
import {
  XLSX,
  type XlsxReadOptions,
  type XlsxWriteOptions,
  type IParseStream,
  type IWritableStream
} from "@excel/xlsx/xlsx.browser";

/** Get (or lazily create) the xlsx IO handle bound to a workbook. */
export function getXlsxIo(wb: WorkbookData): XLSX {
  if (!wb._xlsx) {
    wb._xlsx = new XLSX(wb);
  }
  return wb._xlsx;
}

// =============================================================================
// Cross-platform flat IO functions (the canonical public surface).
// =============================================================================

/** Serialize a workbook to xlsx bytes. */
export function toXlsxBuffer(wb: WorkbookData, options?: XlsxWriteOptions): Promise<Uint8Array> {
  return getXlsxIo(wb).writeBuffer(options);
}

/** Load xlsx bytes into a workbook (mutates and returns `wb`). */
export function loadXlsx(
  wb: WorkbookData,
  data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
  options?: XlsxReadOptions
): Promise<WorkbookData> {
  return getXlsxIo(wb).load(data, options) as unknown as Promise<WorkbookData>;
}

/** Read a workbook from a parse stream. */
export function readXlsxStream(
  wb: WorkbookData,
  stream: IParseStream,
  options?: XlsxReadOptions
): Promise<WorkbookData> {
  return getXlsxIo(wb).read(stream, options) as unknown as Promise<WorkbookData>;
}

/** Write a workbook to a writable stream. */
export function writeXlsxStream(
  wb: WorkbookData,
  stream: IWritableStream,
  options?: XlsxWriteOptions
): Promise<unknown> {
  return getXlsxIo(wb).write(stream, options);
}
