/**
 * `Range` namespace surface — geometric range helpers plus range reads.
 *
 * `import { Range } from "documonster/excel"` → `Range.create("A1:B2")`,
 * `Range.contains(r, "A1")`, `Range.getValues(ws, "G7:H19")`, `Range.setValues(ws, "A1:B2", rows)`.
 *
 * The subject holds two kinds of member: geometry on a range handle, which touches no sheet, and the
 * `(ws, range, …)` reads and writes of the cells a range covers on a sheet.
 */
import type { CellData } from "@excel/core/cell";
import {
  cellSetAlignment,
  cellSetBorder,
  cellSetFill,
  cellSetFont,
  cellSetNumFmt,
  cellSetProtection,
  cellSetStyle,
  cellSetValue,
  checkCellValue,
  checkSharedFormulaAt
} from "@excel/core/cell";
import type { RangeData, RangeInput } from "@excel/core/range";
import {
  rangeContains,
  rangeContainsEx,
  rangeCount,
  rangeCreate,
  rangeExpand,
  rangeExpandToAddress,
  rangeForEachAddress,
  rangeIntersects,
  rangeToString
} from "@excel/core/range";
import {
  checkWorksheet,
  getCell,
  getRangeValues,
  resolveRangeBounds
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { ExcelInvalidAddressError } from "@excel/errors";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import { checkStreamRowsHeld, checkStreamRowsOpen } from "@excel/surface/stream-targets";
import type {
  Alignment,
  Borders,
  CellValue,
  CellValueInput,
  DecodedAddress,
  Fill,
  Font,
  NumFmt,
  Protection,
  Style
} from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import { colCache } from "@excel/utils/col-cache";
import { checkStyle, checkStyleFacet } from "@excel/utils/style-checks";

/** Refuse something that is not a range (`Range.create`) where one is required. */
function checkRange(where: string, r: unknown, field = "r"): asserts r is RangeData {
  if (typeof (r as { top?: unknown } | null | undefined)?.top !== "number") {
    argumentChecks(where).fail(field, "a range (Range.create)", r);
  }
}

/** Whether the range covers the cell at `addressStr` (`"B2"`). */
export function contains(r: RangeData, addressStr: string): boolean {
  checkRange("Range.contains", r);
  return rangeContains(r, addressStr);
}

/**
 * Whether the range covers the cell `address`, a decoded address with a 1-based `row` and `col`. For an `"A1"` string
 * use {@link contains}.
 *
 * @throws {ExcelArgumentError} when `address` is not a decoded address — a string was compared as one and always
 * answered `false`.
 */
export function containsCell(r: RangeData, address: DecodedAddress): boolean {
  checkRange("Range.containsCell", r);
  const check = argumentChecks("Range.containsCell");
  check.required("address", address);
  if (check.record("address", address)) {
    check.required("address.row", address.row);
    check.number("address.row", address.row);
    check.required("address.col", address.col);
    check.number("address.col", address.col);
  }
  return rangeContainsEx(r, address);
}

/** Whether the two ranges share a cell (on the same sheet, when both name one). */
export function intersects(r: RangeData, other: RangeData): boolean {
  checkRange("Range.intersects", r);
  checkRange("Range.intersects", other, "other");
  return rangeIntersects(r, other);
}

/** The range as `A1:B2` text. */
export function toString(r: RangeData): string {
  checkRange("Range.toString", r);
  return rangeToString(r);
}

/** Visit every cell of the range, row by row, as `(address, row, col)`. */
export function eachAddress(
  r: RangeData,
  cb: (address: string, row: number, col: number) => void
): void {
  checkRange("Range.eachAddress", r);
  if (typeof cb !== "function") {
    argumentChecks("Range.eachAddress").fail("cb", "a function", cb);
  }
  rangeForEachAddress(r, cb);
}

/** How many cells the range covers. */
export function getCount(r: RangeData): number {
  checkRange("Range.getCount", r);
  return rangeCount(r);
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;

function inGrid(value: number, max: number): boolean {
  return Number.isInteger(value) && value >= 1 && value <= max;
}

function boundsAreValid(top: number, left: number, bottom: number, right: number): boolean {
  return (
    inGrid(top, MAX_ROW) &&
    inGrid(bottom, MAX_ROW) &&
    inGrid(left, MAX_COL) &&
    inGrid(right, MAX_COL)
  );
}

function describeInput(args: readonly unknown[]): string {
  return args.map(arg => (typeof arg === "object" ? JSON.stringify(arg) : String(arg))).join(", ");
}

/**
 * A range — from `"A1:B2"`, two corners (`"A1", "B2"`), four 1-based numbers (`top, left, bottom, right`) or another
 * range. With no arguments it is the empty range that {@link expand} and {@link expandToAddress} grow from.
 *
 * @throws {ExcelInvalidAddressError} when the arguments do not name cells inside A1:XFD1048576 — `"garbage"` used to give
 * an empty object, and `"A0"` a range starting at row 0.
 */
export function create(...args: RangeInput[]): RangeData {
  const r = rangeCreate(...args);
  const empty = r.top === 0 && r.left === 0 && r.bottom === 0 && r.right === 0;
  const fromEmpty =
    args.length === 0 ||
    (args.length === 1 && typeof args[0] === "object" && !Array.isArray(args[0]) && empty);
  if (!(fromEmpty && empty) && !boundsAreValid(r.top, r.left, r.bottom, r.right)) {
    throw new ExcelInvalidAddressError(
      describeInput(args),
      "a range names cells inside A1:XFD1048576",
      {
        field: "args",
        where: "Range.create"
      }
    );
  }
  return r;
}

/** Grow the range to cover the rectangle `top, left, bottom, right` (1-based). */
export function expand(
  r: RangeData,
  top: number,
  left: number,
  bottom: number,
  right: number
): void {
  checkRange("Range.expand", r);
  if (!boundsAreValid(top, left, bottom, right)) {
    const bounds = { top, left, bottom, right };
    const field = (Object.keys(bounds) as (keyof typeof bounds)[]).find(
      key => !inGrid(bounds[key], key === "top" || key === "bottom" ? MAX_ROW : MAX_COL)
    );
    throw new ExcelInvalidAddressError(
      describeInput([top, left, bottom, right]),
      "the bounds are whole numbers inside A1:XFD1048576",
      { field, where: "Range.expand" }
    );
  }
  rangeExpand(r, top, left, bottom, right);
}

/** Grow the range to cover the cell at `addressStr` (`"B2"`). */
export function expandToAddress(r: RangeData, addressStr: string): void {
  checkRange("Range.expandToAddress", r);
  const address = typeof addressStr === "string" ? colCache.decodeEx(addressStr) : undefined;
  if (
    address === undefined ||
    !("row" in address && "col" in address) ||
    !boundsAreValid(address.row, address.col, address.row, address.col)
  ) {
    throw new ExcelInvalidAddressError(
      String(addressStr),
      "a cell reference such as B2, inside A1:XFD1048576",
      { field: "addressStr", where: "Range.expandToAddress" }
    );
  }
  rangeExpandToAddress(r, addressStr);
}

// --- writing the cells of a range on a sheet ---

type Bounds = { top: number; left: number; bottom: number; right: number };

/** The sheet and the bounds of `range` on it, checked as {@link getValues} checks them. */
function sheetBounds(
  where: string,
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData
): Bounds {
  checkWorksheet(ws);
  if (typeof range !== "string") {
    checkRange(where, range, "range");
  }
  return resolveRangeBounds(ws, range, { where, field: "range" });
}

/**
 * {@link sheetBounds} for a write: refused before any cell is touched when a streaming writer has written a row of the
 * range — rows are written in order, so its top row decides — and the call then changes nothing.
 */
function writableBounds(
  where: string,
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData
): Bounds {
  const bounds = sheetBounds(where, ws, range);
  checkStreamRowsOpen(ws, where, bounds.top);
  return bounds;
}

/**
 * The values of the cells a range covers, row-major: always `bottom - top + 1` rows of `right - left + 1` values, a
 * cell that does not exist reading as `null`. Reading creates nothing. Formula cells yield their formula record, dates
 * a `Date`, rich text and hyperlinks their object forms; every cell of a merged region yields the master's value.
 *
 * On a streaming worksheet writer the range must lie in rows it still holds: one it has written is refused, rather
 * than read as empty.
 *
 * @param range - An A1 range (`"G7:H19"`), a single address, or a {@link RangeData} handle. Reversed corners are
 * normalised; a sheet qualifier must name `ws`.
 * @throws {ExcelInvalidAddressError} when the range cannot be decoded, is unbounded (`"A:A"`) or unset.
 * @throws {ExcelStreamStateError} when a streaming writer has written a row of the range.
 */
export function getValues(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData
): CellValue[][] {
  const { top } = sheetBounds("Range.getValues", ws, range);
  checkStreamRowsHeld(ws, "Range.getValues", top);
  return getRangeValues(ws as WorksheetData, range);
}

/**
 * Write a matrix of values into a range, its first row and column at the range's top-left cell.
 *
 * `values` is row-major, as {@link getValues} returns it. Its rows may be shorter than the range and may differ in
 * length; a cell no row reaches is left as it is, so `Range.setValues(ws, "A1:C3", [[1]])` writes A1 alone. Every
 * index a row does reach is written, `undefined` and an array hole included, which clear the cell as
 * `Cell.setValue(ws, addr, undefined)` does. Each value is what `Cell.setValue` takes and is checked as it checks it; a
 * shared formula's master must already be on the sheet.
 *
 * Every value is checked before any cell is written, so a call that throws leaves the sheet as it was.
 *
 * @param range - An A1 range (`"A1:C3"`), a single address, or a {@link RangeData} handle (`Range.create`).
 * @throws {ExcelArgumentError} when `values` is not an array of arrays, has more rows or a row more values than the
 * range, or holds a value `Cell.setValue` refuses.
 * @throws {ExcelInvalidAddressError} when the range is malformed, unbounded or unset.
 *
 * @example
 * ```ts
 * import { Range } from "documonster/excel";
 *
 * Range.setValues(ws, "A1:C2", [
 *   ["Name", "Qty", "Price"],
 *   ["Apple", 3, 1.5]
 * ]);
 * ```
 */
export function setValues(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  values: readonly (readonly CellValueInput[])[]
): void {
  const where = "Range.setValues";
  const { top, left, bottom, right } = writableBounds(where, ws, range);
  const check = argumentChecks(where);
  if (!Array.isArray(values)) {
    check.fail("values", "an array of rows", values);
  }
  const height = bottom - top + 1;
  const width = right - left + 1;
  if (values.length > height) {
    check.fail("values", `at most ${height} row(s) for the range`, `${values.length} rows`);
  }
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!Array.isArray(row)) {
      check.fail(`values[${i}]`, "an array of values", row);
    }
    if (row.length > width) {
      check.fail(`values[${i}]`, `at most ${width} value(s) for the range`, `${row.length} values`);
    }
    for (let j = 0; j < row.length; j++) {
      const site = { where, field: `values[${i}][${j}]` };
      checkCellValue(row[j], site);
      checkSharedFormulaAt(ws as WorksheetData, top + i, left + j, row[j], site);
    }
  }
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    for (let j = 0; j < row.length; j++) {
      cellSetValue(getCell(ws as WorksheetData, top + i, left + j), row[j]);
    }
  }
}

/** Apply `apply` to every cell of the range, creating the cells it covers. */
function eachCellOf(
  bounds: Bounds,
  ws: WorksheetData | StreamWorksheet,
  apply: (cell: CellData) => void
): void {
  for (let row = bounds.top; row <= bounds.bottom; row++) {
    for (let col = bounds.left; col <= bounds.right; col++) {
      apply(getCell(ws as WorksheetData, row, col));
    }
  }
}

/**
 * Merge a partial style into every cell of the range, exactly as `Cell.setStyle` does into one: the facets `style`
 * carries replace the cell's, and the others are kept.
 *
 * The cells share the facet objects, as cells given one object through `Cell.setStyle` do; editing one cell's style
 * through `Cell.getStyle` gives that cell its own copy first. `style` is checked before any cell is touched.
 *
 * @example
 * ```ts
 * import { Range } from "documonster/excel";
 *
 * Range.setStyle(ws, "A1:C1", { font: { bold: true }, numFmt: "0.00" });
 * ```
 */
export function setStyle(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  style: Partial<Style>
): void {
  const bounds = writableBounds("Range.setStyle", ws, range);
  checkStyle("Range.setStyle", style);
  eachCellOf(bounds, ws, cell => cellSetStyle(cell, style));
}

/**
 * Set the font of every cell of the range, as `Cell.setFont` sets one cell's: the font is replaced, and `undefined`
 * clears it.
 */
export function setFont(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  value: Partial<Font> | undefined
): void {
  const bounds = writableBounds("Range.setFont", ws, range);
  checkStyleFacet("Range.setFont", "font", value, "value");
  eachCellOf(bounds, ws, cell => cellSetFont(cell, value));
}

/** Set the fill of every cell of the range, as `Cell.setFill` does; `undefined` clears it. */
export function setFill(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  value: Fill | undefined
): void {
  const bounds = writableBounds("Range.setFill", ws, range);
  checkStyleFacet("Range.setFill", "fill", value, "value");
  eachCellOf(bounds, ws, cell => cellSetFill(cell, value));
}

/**
 * Set the borders of every cell of the range, as `Cell.setBorder` does; `undefined` clears them.
 *
 * Each cell gets `value` whole — `{ top: … }` puts a top border on every cell of the range, not on its top edge. An
 * outline is the per-edge calls on the edge rows and columns (`Range.setBorder(ws, "A1:C1", { top })`, …).
 */
export function setBorder(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  value: Partial<Borders> | undefined
): void {
  const bounds = writableBounds("Range.setBorder", ws, range);
  checkStyleFacet("Range.setBorder", "border", value, "value");
  eachCellOf(bounds, ws, cell => cellSetBorder(cell, value));
}

/** Set the alignment of every cell of the range, as `Cell.setAlignment` does; `undefined` clears it. */
export function setAlignment(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  value: Partial<Alignment> | undefined
): void {
  const bounds = writableBounds("Range.setAlignment", ws, range);
  checkStyleFacet("Range.setAlignment", "alignment", value, "value");
  eachCellOf(bounds, ws, cell => cellSetAlignment(cell, value));
}

/** Set the number format of every cell of the range, as `Cell.setNumFmt` does; `undefined` clears it. */
export function setNumFmt(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  value: string | NumFmt | undefined
): void {
  const bounds = writableBounds("Range.setNumFmt", ws, range);
  checkStyleFacet("Range.setNumFmt", "numFmt", value, "value");
  eachCellOf(bounds, ws, cell => cellSetNumFmt(cell, value));
}

/** Set the protection of every cell of the range, as `Cell.setProtection` does; `undefined` clears it. */
export function setProtection(
  ws: WorksheetData | StreamWorksheet,
  range: string | RangeData,
  value: Partial<Protection> | undefined
): void {
  const bounds = writableBounds("Range.setProtection", ws, range);
  checkStyleFacet("Range.setProtection", "protection", value, "value");
  eachCellOf(bounds, ws, cell => cellSetProtection(cell, value));
}

/** A range handle. */
export type { RangeData as Handle } from "@excel/core/range";

// --- deprecated names ---

/** @deprecated Use {@link eachAddress}. Removed in 1.0. */
export const forEachAddress = eachAddress;
/** @deprecated Use {@link getCount}. Removed in 1.0. */
export const count = getCount;
