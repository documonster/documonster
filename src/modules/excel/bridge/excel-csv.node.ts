/**
 * `ExcelCsv` namespace surface — Node entry.
 *
 * The cross-platform members plus the file-path variants, with `toStream` and
 * `createImportStream` typed as the `stream.Readable` / `stream.Writable` Node actually returns.
 */
export {
  readCsvWorkbook as read,
  readCsv as addWorksheet,
  writeCsvString as format,
  writeCsvBuffer as toBuffer,
  writeCsvBufferSync as toBufferSync,
  writeCsvStream as writeStream
} from "@excel/bridge/csv-bridge";
export {
  readCsvFileWorkbook as readFile,
  writeCsvFile as writeFile,
  createCsvReadStream as toStream,
  createCsvImportStream as createImportStream
} from "@excel/bridge/csv-bridge.node";
