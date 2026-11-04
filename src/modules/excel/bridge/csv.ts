/**
 * `documonster/excel/csv` — CSV ↔ Workbook.
 *
 * Tree-shakeable CSV import/export for the excel Workbook. Importing this
 * subpath is what pulls the CSV parser/formatter into a bundle; the core
 * `Workbook` never references it.
 */

import {
  createCsvReadStream as createCsvReadStreamImpl,
  createCsvWriteStream as createCsvWriteStreamImpl,
  readCsv as readCsvImpl,
  writeCsv as writeCsvImpl,
  writeCsvBuffer as writeCsvBufferImpl
} from "@excel/bridge/csv-bridge";

export * as ExcelCsv from "@excel/bridge/excel-csv";

// The errors this entry's functions throw — the same class objects `documonster/excel` and `documonster/csv`
// export, so `instanceof` and the guards agree whichever entry a caller imports them from. An argument the bridge
// cannot use is an `ExcelArgumentError`; text the CSV parser cannot read keeps its `CsvError`.
export {
  ExcelArgumentError,
  ExcelDownloadError,
  ExcelError,
  ExcelFileError,
  InvalidValueTypeError,
  isExcelError,
  WorksheetNameError
} from "@excel/errors";
export type { ExcelFileErrorOptions, InvalidValueTypeErrorOptions } from "@excel/errors";
export { CsvError, isCsvError } from "@csv/errors";
export type { BaseErrorOptions } from "@utils/errors";
export type {
  CsvInput,
  CsvOptions,
  CsvReadOptions,
  CsvWriteOptions
} from "@excel/bridge/csv-bridge";

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
