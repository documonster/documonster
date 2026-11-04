/**
 * `documonster/excel/markdown` — Node entry (adds the file-path variants).
 */

import {
  readMarkdown as readMarkdownImpl,
  readMarkdownAll as readMarkdownAllImpl,
  writeMarkdown as writeMarkdownImpl,
  writeMarkdownBuffer as writeMarkdownBufferImpl
} from "@excel/bridge/markdown-bridge";
import {
  readMarkdownAllFile as readMarkdownAllFileImpl,
  readMarkdownFile as readMarkdownFileImpl,
  writeMarkdownFile as writeMarkdownFileImpl
} from "@excel/bridge/markdown-bridge.node";

export type {
  MarkdownInput,
  MarkdownReadOptions,
  MarkdownWriteOptions
} from "@excel/bridge/markdown-bridge";
export * as ExcelMarkdown from "@excel/bridge/excel-markdown.node";

// The errors this entry's functions throw — the same class objects `documonster/excel` and `documonster/markdown`
// export, so `instanceof` and the guards agree whichever entry a caller imports them from. An argument the bridge
// cannot use is an `ExcelArgumentError`; a document with no table the parser can read keeps its `MarkdownParseError`.
export {
  ExcelArgumentError,
  ExcelError,
  ExcelFileError,
  InvalidValueTypeError,
  isExcelError,
  WorksheetNameError
} from "@excel/errors";
export type { ExcelFileErrorOptions, InvalidValueTypeErrorOptions } from "@excel/errors";
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
/** @deprecated Use {@link ExcelMarkdown.readFile} for a new workbook, or {@link ExcelMarkdown.addWorksheet}. Removed in 1.0. */
export const readMarkdownFile = readMarkdownFileImpl;
/**
 * @deprecated Use `ExcelMarkdown.readFile(path, { tables: "all" })` for a new workbook, or
 * {@link ExcelMarkdown.addWorksheets}. Removed in 1.0.
 */
export const readMarkdownAllFile = readMarkdownAllFileImpl;
/** @deprecated Use {@link ExcelMarkdown.writeFile}. Removed in 1.0. */
export const writeMarkdownFile = writeMarkdownFileImpl;
