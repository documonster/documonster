/**
 * `Range` namespace surface — geometric range helpers plus range reads.
 *
 * `import { Range } from "documonster/excel"` → `Range.create("A1:B2")`,
 * `Range.contains(r, "A1")`, `Range.getValues(ws, "G7:H19")`.
 */
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
import { ExcelInvalidAddressError } from "@excel/errors";
import type { DecodedAddress } from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import { colCache } from "@excel/utils/col-cache";

export { getRangeValues as getValues } from "@excel/core/worksheet-core";

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

/** A range handle. */
export type { RangeData as Handle } from "@excel/core/range";

// --- deprecated names ---

/** @deprecated Use {@link eachAddress}. Removed in 1.0. */
export const forEachAddress = eachAddress;
/** @deprecated Use {@link getCount}. Removed in 1.0. */
export const count = getCount;
