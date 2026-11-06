/**
 * `Workbook` namespace surface — Node entry.
 *
 * `import { Workbook } from "documonster/excel"` → `Workbook.create()`,
 * `Workbook.addWorksheet(wb, name)`, `Workbook.toBuffer(wb)`, `Workbook.read(bytes)`,
 * `Workbook.readFile(path)`, `Workbook.writeFile(wb, path)`, …
 *
 * Re-exports the workbook management functions (de-prefixed) plus the
 * Node xlsx IO (includes file-path `readFile` / `writeFile`).
 */
import {
  createStreamWriter as createStreamWriterImpl,
  createStreamReader as createStreamReaderImpl,
  toBuffer as toBufferImpl,
  toStream as toStreamImpl,
  writeFile as writeFileImpl,
  writeStream as writeStreamImpl
} from "@excel/core/workbook";
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
  getDefaultFont,
  setDefaultFont,
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

export { read, readWithDiagnostics, readFile, readStream } from "@excel/core/workbook";
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
export const writeFile = /* @__PURE__ */ workbookOnly(
  "Workbook.writeFile",
  WRITER_FINISHES,
  writeFileImpl
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
