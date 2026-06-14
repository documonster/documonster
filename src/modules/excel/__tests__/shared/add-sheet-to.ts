import { addWorksheet } from "@excel/workbook";
import { getCell, mergeCells } from "@excel/worksheet";

/**
 * Add a worksheet to either a plain `WorkbookData` (flat `addWorksheet`) or a
 * streaming `WorkbookWriter` (instance method). Test sheet helpers reuse this.
 */
export function addSheetTo(wb: any, name?: string, options?: any): any {
  return typeof wb.addWorksheet === "function"
    ? wb.addWorksheet(name, options)
    : addWorksheet(wb, name, options);
}

/** Resolve a cell by address on a record worksheet (flat) or streaming writer (method). */
export function cellAt(ws: any, addr: string | number, col?: number): any {
  return typeof ws.getCell === "function" ? ws.getCell(addr, col) : getCell(ws, addr, col);
}

/** Merge cells on a record worksheet (flat) or streaming writer (method). */
export function mergeAt(ws: any, ...cells: any[]): void {
  if (typeof ws.mergeCells === "function") {
    ws.mergeCells(...cells);
  } else {
    mergeCells(ws, ...cells);
  }
}
