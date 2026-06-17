/**
 * Workbook - Node.js entry point.
 *
 * Re-exports the platform-independent workbook surface (record + flat
 * functions) from `workbook.browser`, and adds the Node-only XLSX file-path
 * IO free functions. The browser/Node split keeps Node-only `fs` and stream
 * code out of browser bundles.
 */

import {
  WorkbookReader,
  type WorkbookReaderOptions,
  type NodeInput
} from "@excel/stream/workbook-reader";
import type { WorkbookReader as WorkbookReaderBrowser } from "@excel/stream/workbook-reader.browser";
import { WorkbookWriter, type WorkbookWriterOptions } from "@excel/stream/workbook-writer";
import type { WorkbookWriter as WorkbookWriterBrowser } from "@excel/stream/workbook-writer.browser";
import type { WorkbookData } from "@excel/workbook.browser";
import { XLSX } from "@excel/xlsx/xlsx";
import type { XlsxReadOptions, XlsxWriteOptions } from "@excel/xlsx/xlsx.browser";

export * from "@excel/workbook.browser";

/**
 * Node-only xlsx IO handle for a workbook. Exposes file-path operations
 * (`readFile` / `writeFile`) plus streaming `read` / `write` that the browser
 * variant omits. Free function so it tree-shakes; only pulled into Node bundles.
 */
export function getXlsxIo(wb: WorkbookData): XLSX {
  const slot = wb as WorkbookData & { _xlsxNode?: XLSX };
  if (!slot._xlsxNode) {
    slot._xlsxNode = new XLSX(wb);
  }
  return slot._xlsxNode;
}

/** Node streaming workbook writer factory (accepts `{ filename }`). */
export function createStreamWriter(options?: WorkbookWriterOptions): WorkbookWriterBrowser {
  return new WorkbookWriter(options) as unknown as WorkbookWriterBrowser;
}

/** Node streaming workbook reader factory (accepts a file-path string). */
export function createStreamReader(
  input: NodeInput,
  options?: WorkbookReaderOptions
): WorkbookReaderBrowser {
  return new WorkbookReader(input, options) as unknown as WorkbookReaderBrowser;
}

// =============================================================================
// Cross-platform + Node xlsx IO (flat functions). The Node variant binds the
// file-capable XLSX (`readFile` / `writeFile`).
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

/** Node-only: read a workbook from an xlsx file path (mutates and returns `wb`). */
export function readXlsxFile(
  wb: WorkbookData,
  filename: string,
  options?: XlsxReadOptions
): Promise<WorkbookData> {
  return getXlsxIo(wb).readFile(filename, options) as unknown as Promise<WorkbookData>;
}

/** Node-only: write a workbook to an xlsx file path. */
export function writeXlsx(
  wb: WorkbookData,
  filename: string,
  options?: XlsxWriteOptions
): Promise<void> {
  return getXlsxIo(wb).writeFile(filename, options);
}

/** Read a workbook from a parse stream (mutates and returns `wb`). */
export function readXlsxStream(
  wb: WorkbookData,
  stream: unknown,
  options?: XlsxReadOptions
): Promise<WorkbookData> {
  return getXlsxIo(wb).read(stream as never, options) as unknown as Promise<WorkbookData>;
}

/** Write a workbook to a writable stream. */
export function writeXlsxStream(
  wb: WorkbookData,
  stream: unknown,
  options?: XlsxWriteOptions
): Promise<unknown> {
  return getXlsxIo(wb).write(stream as never, options);
}

export type { CsvOptions, CsvInput } from "@excel/bridge/csv-bridge";
export type {
  WorkbookModel,
  WorkbookMedia,
  WorkbookProtectionModel,
  ExternalLinkModel,
  ExternalLinkCachedSheet
} from "@excel/workbook.browser";
export type {
  AddChartsheetOptions,
  AddPivotChartsheetOptions,
  ChartsheetOptions,
  ChartsheetViewOptions
} from "@excel/chartsheet";
