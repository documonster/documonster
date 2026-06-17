/**
 * `Stream` namespace surface — Node entry.
 *
 * The streaming API is a class-based, incremental, handle-oriented paradigm
 * (distinct from the random-access document model). `Stream` bundles the
 * streaming writer/reader classes together with the handle-level operations
 * used on the `CellData` / `RowData` returned by a streaming worksheet
 * writer, so streaming code never has to reach for the internal flat helpers.
 *
 * `import { Stream } from "documonster/excel"` →
 *   `const wb = new Stream.WorkbookWriter({ filename });`
 *   `const ws = wb.addWorksheet("Sheet1");`
 *   `const row = ws.addRow([1, 2, 3]);`
 *   `Stream.setCellValue(ws.getCell("A1"), 42);`
 *   `Stream.setRowFont(row, { bold: true });`
 *   `Stream.commitRow(row);`
 */
export { WorkbookWriter } from "@excel/stream/workbook-writer";
export { WorkbookReader } from "@excel/stream/workbook-reader";
export type { WorkbookWriterOptions } from "@excel/stream/workbook-writer.browser";
export type { WorkbookReaderOptions } from "@excel/stream/workbook-reader.browser";

// --- streaming cell handle operations (operate on a `CellData`) ---
export {
  cellGetValue as getCellValue,
  cellSetValue as setCellValue,
  cellSetFont as setCellFont,
  cellSetFill as setCellFill,
  cellSetBorder as setCellBorder,
  cellSetAlignment as setCellAlignment,
  cellSetNumFmt as setCellNumFmt,
  cellSetNote as setCellNote,
  cellSetComment as setCellComment
} from "@excel/cell";

// --- streaming row handle operations (operate on a `RowData`) ---
export {
  rowValues,
  rowSetFont as setRowFont,
  rowSetFill as setRowFill,
  rowSetBorder as setRowBorder,
  rowSetAlignment as setRowAlignment
} from "@excel/row";
export { rowGetCell as rowCell, rowCommit as commitRow } from "@excel/worksheet-core";

/** A streaming cell handle. */
export type { CellData as CellHandle } from "@excel/cell";
/** A streaming row handle. */
export type { RowData as RowHandle } from "@excel/row";
