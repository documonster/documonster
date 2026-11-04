/**
 * `Address` namespace surface — stateless A1-notation encode/decode helpers.
 *
 * `import { Address } from "documonster/excel"` →
 *   `Address.decodeCell("B3")`, `Address.encodeCol(2)`,
 *   `Address.decodeRange("A1:C5")`, `Address.quoteSheetName("My Sheet")`.
 *
 * The coordinates are **0-based** (column A and row 1 are `0`), the convention of the spreadsheet libraries these
 * helpers interoperate with — unlike the `Cell`, `Row` and `Column` namespaces, which number from 1. Each one refuses
 * a coordinate or a reference outside the grid instead of answering with `undefined`, `NaN` or a negative row.
 */
import { ColumnOutOfBoundsError, InvalidAddressError, RowOutOfBoundsError } from "@excel/errors";
import type { CellAddress, SheetRange } from "@excel/utils/address";
import {
  decodeCell as decodeCellImpl,
  decodeCol as decodeColImpl,
  decodeRow as decodeRowImpl,
  encodeCell as encodeCellImpl,
  encodeCol as encodeColImpl,
  encodeRange as encodeRangeImpl,
  encodeRow as encodeRowImpl,
  quoteSheetName as quoteSheetNameImpl
} from "@excel/utils/address";
import { argumentChecks } from "@excel/utils/arguments";

/**
 * `sheetName` as a formula writes it — quoted, with `'` doubled, when it is not a plain identifier or reads as a
 * reference or a boolean (`'My Sheet'`, `'A1'`, `'TRUE'`).
 *
 * @throws {ExcelArgumentError} when `sheetName` is not a string.
 */
export function quoteSheetName(sheetName: string): string {
  argumentChecks("Address.quoteSheetName").required("sheetName", sheetName);
  argumentChecks("Address.quoteSheetName").string("sheetName", sheetName);
  return quoteSheetNameImpl(sheetName);
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;

function checkCol(where: string, field: string, col: number): void {
  if (typeof col !== "number") {
    argumentChecks(where).fail(field, "a 0-based column number", col);
  }
  if (!Number.isInteger(col) || col < 0 || col >= MAX_COL) {
    throw new ColumnOutOfBoundsError(col, "a 0-based column is a whole number from 0 to 16383", {
      field,
      where
    });
  }
}

function checkRow(where: string, field: string, row: number): void {
  if (typeof row !== "number") {
    argumentChecks(where).fail(field, "a 0-based row number", row);
  }
  if (!Number.isInteger(row) || row < 0 || row >= MAX_ROW) {
    throw new RowOutOfBoundsError(row, "a 0-based row is a whole number from 0 to 1048575", {
      field,
      where
    });
  }
}

/**
 * Decode column letters to a 0-based number.
 * @example decodeCol("A") // => 0
 * @example decodeCol("AA") // => 26
 */
export function decodeCol(colstr: string): number {
  if (typeof colstr !== "string" || !/^[A-Za-z]{1,3}$/.test(colstr)) {
    throw new InvalidAddressError(String(colstr), "a column is one to three letters, A to XFD", {
      field: "colstr",
      where: "Address.decodeCol"
    });
  }
  return decodeColImpl(colstr);
}

/**
 * Encode a 0-based column number as letters.
 * @example encodeCol(0) // => "A"
 * @example encodeCol(26) // => "AA"
 */
export function encodeCol(col: number): string {
  checkCol("Address.encodeCol", "col", col);
  return encodeColImpl(col);
}

/**
 * Decode a row number string to a 0-based number.
 * @example decodeRow("1") // => 0
 */
export function decodeRow(rowstr: string): number {
  const row = typeof rowstr === "string" && /^\d+$/.test(rowstr) ? decodeRowImpl(rowstr) : NaN;
  if (!Number.isInteger(row) || row < 0 || row >= MAX_ROW) {
    throw new InvalidAddressError(String(rowstr), "a row is a number from 1 to 1048576", {
      field: "rowstr",
      where: "Address.decodeRow"
    });
  }
  return row;
}

/**
 * Encode a 0-based row number as its row string.
 * @example encodeRow(0) // => "1"
 */
export function encodeRow(row: number): string {
  checkRow("Address.encodeRow", "row", row);
  return encodeRowImpl(row);
}

/**
 * Decode a cell reference (`"B2"`, `"$B$2"`) to 0-based `{ c, r }`.
 * @example decodeCell("B2") // => { c: 1, r: 1 }
 */
export function decodeCell(cstr: string): CellAddress {
  return decodeCellAt("Address.decodeCell", "cstr", cstr);
}

function decodeCellAt(where: string, field: string, cstr: string): CellAddress {
  // The shape first, so anything that is not a reference is an `InvalidAddressError`, whatever its letters; a
  // well-formed reference past XFD is then the decoder's `ColumnOutOfBoundsError`.
  const cell =
    typeof cstr === "string" && /^\$?[A-Za-z]+\$?\d+$/.test(cstr)
      ? decodeCellImpl(cstr)
      : { c: NaN, r: NaN };
  if (
    !Number.isInteger(cell.c) ||
    !Number.isInteger(cell.r) ||
    cell.c < 0 ||
    cell.c >= MAX_COL ||
    cell.r < 0 ||
    cell.r >= MAX_ROW
  ) {
    throw new InvalidAddressError(
      String(cstr),
      "a cell reference is column letters then a row number, inside A1:XFD1048576",
      { field, where }
    );
  }
  return cell;
}

/**
 * Encode 0-based `{ c, r }` as a cell reference.
 * @example encodeCell({ c: 1, r: 1 }) // => "B2"
 */
export function encodeCell(cell: CellAddress): string {
  return encodeCellAt("Address.encodeCell", "cell", cell);
}

function encodeCellAt(where: string, field: string, cell: CellAddress): string {
  const check = argumentChecks(where);
  check.required(field, cell);
  check.record(field, cell);
  checkCol(where, `${field}.c`, cell.c);
  checkRow(where, `${field}.r`, cell.r);
  return encodeCellImpl(cell);
}

/**
 * Decode a range (`"A1:B2"`, or a single cell) to 0-based `{ s, e }`.
 * @example decodeRange("A1:B2") // => { s: { c: 0, r: 0 }, e: { c: 1, r: 1 } }
 */
export function decodeRange(range: string): SheetRange {
  const where = "Address.decodeRange";
  const idx = typeof range === "string" ? range.indexOf(":") : -1;
  if (idx === -1) {
    const cell = decodeCellAt(where, "range", range);
    return { s: cell, e: { ...cell } };
  }
  return {
    s: decodeCellAt(where, "range", range.slice(0, idx)),
    e: decodeCellAt(where, "range", range.slice(idx + 1))
  };
}

/** Encode a 0-based range as `"A1:B2"` — or `"A1"` when it is one cell. */
export function encodeRange(range: SheetRange): string;
/** Encode the 0-based range from `start` to `end`. */
export function encodeRange(start: CellAddress, end: CellAddress): string;
export function encodeRange(startOrRange: CellAddress | SheetRange, end?: CellAddress): string {
  const where = "Address.encodeRange";
  const check = argumentChecks(where);
  const whole = end === undefined;
  check.required(whole ? "range" : "start", startOrRange);
  check.record(whole ? "range" : "start", startOrRange);
  const [s, e] = whole
    ? [(startOrRange as SheetRange).s, (startOrRange as SheetRange).e]
    : [startOrRange as CellAddress, end];
  encodeCellAt(where, whole ? "range.s" : "start", s);
  encodeCellAt(where, whole ? "range.e" : "end", e);
  return encodeRangeImpl(s, e);
}
