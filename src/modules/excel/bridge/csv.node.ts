/**
 * `documonster/excel/csv` — Node entry.
 *
 * Re-exports the cross-platform CSV functions plus the Node-only file-path
 * variants (`readCsvFile` / `writeCsvFile`).
 */

export {
  readCsv,
  writeCsv,
  writeCsvBuffer,
  createCsvReadStream,
  createCsvWriteStream,
  type CsvInput,
  type CsvOptions
} from "./csv-bridge";
export { readCsvFile, writeCsvFile } from "./csv-bridge.node";
