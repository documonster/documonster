import {
  addCellStyle as addCellStyleImpl,
  createStreamReader as createStreamReaderImpl,
  createStreamWriter as createStreamWriterImpl,
  eachWorksheet as eachWorksheetImpl,
  getChartsheets as listChartsheetsImpl,
  getWorksheets as listWorksheetsImpl,
  importWorksheet as importWorksheetImpl,
  renameChartsheet as renameChartsheetImpl,
  useBuiltInCellStyle as useBuiltInCellStyleImpl
} from "@excel/core/workbook.browser";
/**
 * `Workbook` namespace surface — browser entry.
 *
 * Same as the Node `surface/workbook.ts` but IO comes from
 * `@excel/xlsx-io` (cross-platform `toBuffer` / `read` / streaming only — no
 * Node file-path `readFile` / `writeFile`). The `.browser` same-name swap
 * selects the browser xlsx binding.
 */
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
  readStream,
  writeStream
} from "@excel/core/xlsx-io";
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
