/**
 * `ExcelCsv` namespace surface — CSV ↔ Workbook (cross-platform).
 *
 * `import { ExcelCsv } from "documonster/excel/csv"` →
 *   in: `ExcelCsv.read(input)` (a new workbook from text, bytes, URL, Blob or a readable stream),
 *   `ExcelCsv.addWorksheet(wb, input)` (a new worksheet on `wb`),
 *   `ExcelCsv.createImportStream(wb)` (a writable you write CSV into);
 *   out: `ExcelCsv.format(wb)` (string), `ExcelCsv.toBuffer(wb)` /
 *   `ExcelCsv.toBufferSync(wb)` (bytes), `ExcelCsv.toStream(wb)` (a readable),
 *   `ExcelCsv.writeStream(wb, out)` (into a writable you own).
 *
 * Named `ExcelCsv` rather than `Csv` so it can be imported next to `Csv` from
 * `documonster/csv` without aliasing.
 */
export {
  readCsvWorkbook as read,
  readCsv as addWorksheet,
  writeCsvString as format,
  writeCsvBuffer as toBuffer,
  writeCsvBufferSync as toBufferSync,
  createCsvReadStream as toStream,
  writeCsvStream as writeStream,
  createCsvWriteStream as createImportStream
} from "@excel/bridge/csv-bridge";
