/**
 * `documonster/excel/markdown` — Markdown tables ↔ Workbook.
 *
 * Tree-shakeable Markdown-table import/export for the excel Workbook.
 */

import {
  readMarkdown as readMarkdownImpl,
  readMarkdownAll as readMarkdownAllImpl,
  writeMarkdown as writeMarkdownImpl,
  writeMarkdownBuffer as writeMarkdownBufferImpl
} from "@excel/bridge/markdown-bridge";

export type {
  ExcelMarkdownInput,
  ExcelMarkdownReadOptions,
  ExcelMarkdownWriteOptions
} from "@excel/bridge/markdown-bridge";
export * as ExcelMarkdown from "@excel/bridge/excel-markdown";

// The errors this entry's functions throw — the same class objects `documonster/excel` and `documonster/markdown`
// export, so `instanceof` and the guards agree whichever entry a caller imports them from. An argument the bridge
// cannot use is an `ExcelArgumentError`; a document with no table the parser can read keeps its `MarkdownParseError`.
export {
  ExcelArgumentError,
  ExcelError,
  ExcelFileError,
  ExcelInvalidValueTypeError,
  isExcelError,
  ExcelWorksheetNameError,
  // Deprecated pre-prefix names; removed in 1.0.
  InvalidValueTypeError,
  WorksheetNameError
} from "@excel/errors";
export type {
  ExcelFileErrorOptions,
  ExcelInvalidValueTypeErrorOptions,
  // Deprecated pre-prefix name; removed in 1.0.
  InvalidValueTypeErrorOptions
} from "@excel/errors";
export { MarkdownError, MarkdownParseError, isMarkdownError } from "@markdown/errors";
export type { BaseErrorOptions } from "@utils/errors";

/** @deprecated Use {@link ExcelMarkdown.addWorksheet}. Removed in 1.0. */
export const readMarkdown = readMarkdownImpl;
/** @deprecated Use {@link ExcelMarkdown.addWorksheets}. Removed in 1.0. */
export const readMarkdownAll = readMarkdownAllImpl;
/** @deprecated Use {@link ExcelMarkdown.format}. Removed in 1.0. */
export const writeMarkdown = writeMarkdownImpl;
/**
 * @deprecated Use {@link ExcelMarkdown.toBufferSync} (or the asynchronous
 * {@link ExcelMarkdown.toBuffer}). Removed in 1.0.
 */
export const writeMarkdownBuffer = writeMarkdownBufferImpl;
