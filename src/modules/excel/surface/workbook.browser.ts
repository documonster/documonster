/**
 * `Workbook` namespace surface — browser entry.
 *
 * Same as the Node `surface/workbook.ts` but IO comes from
 * `@excel/xlsx-io` (cross-platform `toBuffer` / `read` / streaming only — no
 * Node file-path `readFile` / `writeFile`). The `.browser` same-name swap
 * selects the browser xlsx binding.
 */
import {
  createStreamReader as createStreamReaderImpl,
  createStreamWriter as createStreamWriterImpl
} from "@excel/core/workbook.browser";
import {
  toBuffer as toBufferImpl,
  toStream as toStreamImpl,
  writeStream as writeStreamImpl
} from "@excel/core/xlsx-io";
import type {
  XlsxReadable as XlsxReadableType,
  XlsxWritable as XlsxWritableType
} from "@excel/core/xlsx-io-types";
import type { XlsxStreamOptions as XlsxStreamOptionsType } from "@excel/core/xlsx-stream";
import { workbookOnly } from "@excel/surface/stream-targets";

export {
  addWorksheet,
  protect,
  unprotect,
  verifyPassword,
  addExternalLink,
  registerPerson,
  getWorksheet,
  listWorksheets,
  eachWorksheet,
  getDefinedNames,
  addCellStyle,
  getCellStyle,
  listCellStyles,
  removeCellStyle,
  useBuiltInCellStyle,
  removeWorksheet,
  addChartsheet,
  addPivotChartsheet,
  getChartsheet,
  listChartsheets,
  removeChartsheet,
  copyChartsheet,
  replaceChartsheetChart,
  renameChartsheet,
  importWorksheet,
  listUnrecognizedValues,
  discardUnrecognizedValues,
  getExternalLink,
  listExternalLinks,
  listPersons,
  registerFunction,
  unregisterFunction,
  getModel,
  setModel,
  getWorksheets,
  getChartsheets,
  eachSheet,
  importSheet,
  defineCellStyle,
  useBuiltinCellStyle
} from "@excel/surface/workbook-edits";

export { createWorkbook as create } from "@excel/core/workbook.browser";

export { read, readWithDiagnostics, readStream } from "@excel/core/xlsx-io";
// A streaming writer is not written by these: it has written most of its file already, in the format it was made
// with, and the options they take could change neither. It finishes itself, with `commit()` or `toBuffer()`.
const WRITER_FINISHES = "is finished with its own commit() or toBuffer()";

export const toBuffer = /* @__PURE__ */ workbookOnly(
  "Workbook.toBuffer",
  WRITER_FINISHES,
  toBufferImpl
);
export const toStream = /* @__PURE__ */ workbookOnly(
  "Workbook.toStream",
  WRITER_FINISHES,
  toStreamImpl
);
export const writeStream = /* @__PURE__ */ workbookOnly(
  "Workbook.writeStream",
  WRITER_FINISHES,
  writeStreamImpl
);
// Published flat from `documonster/excel`; the namespace copies were a second name for each.
/** @deprecated Use the flat `XlsxReadable` type. Removed in 1.0. */
export type XlsxReadable = XlsxReadableType;
/** @deprecated Use the flat `XlsxWritable` type. Removed in 1.0. */
export type XlsxWritable = XlsxWritableType;
/** @deprecated Use the flat `XlsxStreamOptions` type. Removed in 1.0. */
export type XlsxStreamOptions = XlsxStreamOptionsType;

/** @deprecated Use `new Stream.WorkbookWriter(options)`. Removed in 1.0. */
export const createStreamWriter = createStreamWriterImpl;
/** @deprecated Use `new Stream.WorkbookReader(input, options)`. Removed in 1.0. */
export const createStreamReader = createStreamReaderImpl;

/**
 * A workbook handle. Its document properties (`title`, `creator`, `created`, …), `properties`,
 * `calcProperties` and `views` are plain fields to read and assign; everything else is reached through
 * the namespace functions. Internal state is marked `@deprecated` and leaves the published type in 1.0.
 */
export type { WorkbookData as Handle } from "@excel/core/workbook-core";
