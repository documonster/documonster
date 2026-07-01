/**
 * `Stream` namespace surface — browser entry.
 *
 * Same streaming surface as the Node `surface/stream.ts`, but the
 * writer/reader classes resolve to their browser variants (Web Streams,
 * no Node file-path sinks).
 */
export { WorkbookWriter } from "@excel/stream/workbook-writer.browser";
export { WorkbookReader } from "@excel/stream/workbook-reader.browser";
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
} from "@excel/core/cell";

// --- streaming row handle operations (operate on a `RowData`) ---
export {
  rowValues,
  rowSetFont as setRowFont,
  rowSetFill as setRowFill,
  rowSetBorder as setRowBorder,
  rowSetAlignment as setRowAlignment
} from "@excel/core/row";
export { rowGetCell as rowCell, rowCommit as commitRow } from "@excel/core/worksheet-core";

/** A streaming cell handle. */
export type { CellData as CellHandle } from "@excel/core/cell";
/** A streaming row handle. */
export type { RowData as RowHandle } from "@excel/core/row";
