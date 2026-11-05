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
  createStreamReader as createStreamReaderImpl
} from "@excel/core/workbook";
import {
  addCellStyle as addCellStyleImpl,
  eachWorksheet as eachWorksheetImpl,
  getChartsheets as listChartsheetsImpl,
  getWorksheets as listWorksheetsImpl,
  importWorksheet as importWorksheetImpl,
  renameChartsheet as renameChartsheetImpl,
  useBuiltInCellStyle as useBuiltInCellStyleImpl
} from "@excel/core/workbook.browser";
import type {
  XlsxReadable as XlsxReadableType,
  XlsxWritable as XlsxWritableType
} from "@excel/core/xlsx-io-types";
import type { XlsxStreamOptions as XlsxStreamOptionsType } from "@excel/core/xlsx-stream";

export {
  addWorksheet,
  protect,
  addExternalLink,
  registerPerson
} from "@excel/surface/workbook-edits";

export {
  createWorkbook as create,
  getWorksheet,
  getWorksheets as listWorksheets,
  removeWorksheet,
  discardUnrecognizedValues,
  listUnrecognizedValues,
  addChartsheet,
  addPivotChartsheet,
  getChartsheet,
  getChartsheets as listChartsheets,
  removeChartsheet,
  copyChartsheet,
  replaceChartsheetChart,
  getDefinedNames,
  eachWorksheet,
  importWorksheet,
  unprotectWorkbook as unprotect,
  verifyWorkbookPassword as verifyPassword,
  getExternalLink,
  listExternalLinks,
  listPersons,
  registerFunction,
  unregisterFunction,
  addCellStyle,
  getCellStyle,
  listCellStyles,
  removeCellStyle,
  useBuiltInCellStyle,
  getWorkbookModel as getModel
} from "@excel/core/workbook.browser";

export {
  toBuffer,
  toStream,
  read,
  readWithDiagnostics,
  readFile,
  writeFile,
  readStream,
  writeStream
} from "@excel/core/workbook";
// Published flat from `documonster/excel`; the namespace copies were a second name for each.
/** @deprecated Use the flat `XlsxReadable` type. Removed in 1.0. */
export type XlsxReadable = XlsxReadableType;
/** @deprecated Use the flat `XlsxWritable` type. Removed in 1.0. */
export type XlsxWritable = XlsxWritableType;
/** @deprecated Use the flat `XlsxStreamOptions` type. Removed in 1.0. */
export type XlsxStreamOptions = XlsxStreamOptionsType;

export { setModel } from "@excel/surface/workbook-model";

/** @deprecated Use `new Stream.WorkbookWriter(options)`. Removed in 1.0. */
export const createStreamWriter = createStreamWriterImpl;
/** @deprecated Use `new Stream.WorkbookReader(input, options)`. Removed in 1.0. */
export const createStreamReader = createStreamReaderImpl;

/** @deprecated Use `Chartsheet.setName(Workbook.getChartsheet(wb, nameOrIndex), name)`. Removed in 1.0. */
export const renameChartsheet = renameChartsheetImpl;
/** @deprecated Use {@link listWorksheets}. Removed in 1.0. */
export const getWorksheets = listWorksheetsImpl;
/** @deprecated Use {@link listChartsheets}. Removed in 1.0. */
export const getChartsheets = listChartsheetsImpl;
/** @deprecated Use {@link eachWorksheet}. Removed in 1.0. */
export const eachSheet = eachWorksheetImpl;
/** @deprecated Use {@link importWorksheet}. Removed in 1.0. */
export const importSheet = importWorksheetImpl;
/** @deprecated Use {@link addCellStyle}, which returns (a copy of) the stored entry. Removed in 1.0. */
export const defineCellStyle = addCellStyleImpl;
/** @deprecated Use {@link useBuiltInCellStyle}. Removed in 1.0. */
export const useBuiltinCellStyle = useBuiltInCellStyleImpl;

/**
 * A workbook handle. Its document properties (`title`, `creator`, `created`, …), `properties`,
 * `calcProperties` and `views` are plain fields to read and assign; everything else is reached through
 * the namespace functions. Internal state is marked `@deprecated` and leaves the published type in 1.0.
 */
export type { WorkbookData as Handle } from "@excel/core/workbook-core";
