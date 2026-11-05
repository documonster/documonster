/**
 * `documonster/excel/csv` — Node entry.
 *
 * The cross-platform `ExcelCsv` members plus the Node-only file-path variants
 * (`ExcelCsv.readFile` / `ExcelCsv.writeFile`), with `ExcelCsv.toStream` typed
 * as the Node stream it actually returns.
 */

import type {
  ExcelCsvInput as ExcelCsvInputType,
  ExcelCsvOptions as ExcelCsvOptionsType
} from "@excel/bridge/csv-bridge";
import {
  createCsvWriteStream as createCsvWriteStreamImpl,
  readCsv as readCsvImpl,
  writeCsv as writeCsvImpl,
  writeCsvBuffer as writeCsvBufferImpl
} from "@excel/bridge/csv-bridge";
import {
  createCsvReadStream as createCsvReadStreamImpl,
  readCsvFile as readCsvFileImpl,
  writeCsvFile as writeCsvFileImpl
} from "@excel/bridge/csv-bridge.node";

export * as ExcelCsv from "@excel/bridge/excel-csv.node";

// The errors this entry's functions throw — the same class objects `documonster/excel` and `documonster/csv`
// export, so `instanceof` and the guards agree whichever entry a caller imports them from. An argument the bridge
// cannot use is an `ExcelArgumentError`; text the CSV parser cannot read keeps its `CsvError`.
export {
  ExcelArgumentError,
  ExcelDownloadError,
  ExcelError,
  ExcelFileError,
  ExcelInvalidValueTypeError,
  isExcelError,
  ExcelWorksheetNameError,
  // Deprecated pre-prefix names; removed in 1.0.
  InvalidValueTypeError,
  WorksheetNameError
} from "@excel/errors";
export type { ExcelFileErrorOptions, InvalidValueTypeErrorOptions } from "@excel/errors";
export { CsvError, isCsvError } from "@csv/errors";
export type { BaseErrorOptions } from "@utils/errors";
// Prefixed with the namespace they belong to, so they cannot be read as `documonster/csv`'s own options.
export type {
  ExcelCsvInput,
  ExcelCsvOptions,
  ExcelCsvReadOptions,
  ExcelCsvWriteOptions
} from "@excel/bridge/csv-bridge";
/** @deprecated Use {@link ExcelCsvInput}. Removed in 1.0. */
export type CsvInput = ExcelCsvInputType;
/** @deprecated Use {@link ExcelCsvOptions}. Removed in 1.0. */
export type CsvOptions = ExcelCsvOptionsType;

/** @deprecated Use {@link ExcelCsv.addWorksheet}. Removed in 1.0. */
export const readCsv = readCsvImpl;
/**
 * @deprecated Use {@link ExcelCsv.format} for a string or
 * {@link ExcelCsv.writeStream} for a stream. Removed in 1.0.
 */
export const writeCsv = writeCsvImpl;
/** @deprecated Use {@link ExcelCsv.toBuffer}. Removed in 1.0. */
export const writeCsvBuffer = writeCsvBufferImpl;
/** @deprecated Use {@link ExcelCsv.toStream}. Removed in 1.0. */
export const createCsvReadStream = createCsvReadStreamImpl;
/** @deprecated Use {@link ExcelCsv.createImportStream}. Removed in 1.0. */
export const createCsvWriteStream = createCsvWriteStreamImpl;
/**
 * @deprecated Use {@link ExcelCsv.readFile} for a new workbook, or {@link ExcelCsv.addWorksheet} with a
 * file stream to add the sheet to one you have. Removed in 1.0.
 */
export const readCsvFile = readCsvFileImpl;
/** @deprecated Use {@link ExcelCsv.writeFile}. Removed in 1.0. */
export const writeCsvFile = writeCsvFileImpl;
