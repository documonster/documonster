import type { AnchorData, AnchorInput, AnchorModel } from "@excel/core/anchor";
import {
  anchorCol,
  anchorColWidth,
  anchorCreate,
  anchorModel,
  anchorRow,
  anchorRowHeight,
  anchorSetCol,
  anchorSetModel,
  anchorSetRow
} from "@excel/core/anchor";
import { checkAnchor } from "@excel/core/feature-handles";
import type { Worksheet } from "@excel/core/worksheet";
/**
 * `Anchor` namespace surface — drawing anchor handle (cell coordinate + offset).
 *
 * `import { Anchor } from "documonster/excel"` →
 *   `Anchor.create(ws, "B2")`, `Anchor.getCol(a)`, `Anchor.setRow(a, 4)`,
 *   `Anchor.getModel(a)`, `Anchor.clone(a)`.
 */
import { checkWorksheet, checkWorksheetAttached } from "@excel/core/worksheet-core";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks } from "@excel/utils/arguments";
import { isRecordArgument } from "@utils/object";

export {
  anchorAsInstance as asInstance,
  isAnchorData as isAnchor,
  anchorClone as clone
} from "@excel/core/anchor";

// The sheet's grid, 0-based as an anchor counts it, with room for the edge past the last column and row.
const MAX_COL = 16384;
const MAX_ROW = 1048576;

/**
 * Refuse an edit to the anchor of a drawing on a sheet that was removed — as an edit through the sheet is refused. An
 * anchor bound to no sheet is edited freely.
 */
function checkAnchorEditable(a: AnchorData): void {
  if (a.worksheet !== undefined) {
    checkWorksheetAttached(a.worksheet);
  }
}

function position(check: ArgumentChecks, field: string, value: unknown, max: number): void {
  check.required(field, value);
  check.number(field, value, { min: 0, max });
}

/**
 * Set the anchor's column — a fractional position, the fraction becoming an offset into the column.
 *
 * @throws {ExcelArgumentError} when `value` is not a number from 0 to the sheet's last column. It used to be stored as
 * given and refused only by the drawing that used the anchor.
 */
export function setCol(a: AnchorData, value: number): void {
  checkAnchor(a);
  checkAnchorEditable(a);
  position(argumentChecks("Anchor.setCol"), "value", value, MAX_COL);
  anchorSetCol(a, value);
}

/** Set the anchor's row, as {@link setCol} sets the column. @throws {ExcelArgumentError} for a value off the grid. */
export function setRow(a: AnchorData, value: number): void {
  checkAnchor(a);
  checkAnchorEditable(a);
  position(argumentChecks("Anchor.setRow"), "value", value, MAX_ROW);
  anchorSetRow(a, value);
}

function checkModel(check: ArgumentChecks, field: string, value: unknown): value is AnchorModel {
  check.required(field, value);
  if (!check.record(field, value)) {
    return false;
  }
  check.number(`${field}.nativeCol`, value.nativeCol, { min: 0, max: MAX_COL, integer: true });
  check.number(`${field}.nativeRow`, value.nativeRow, { min: 0, max: MAX_ROW, integer: true });
  check.number(`${field}.nativeColOff`, value.nativeColOff, { min: 0, integer: true });
  check.number(`${field}.nativeRowOff`, value.nativeRowOff, { min: 0, integer: true });
  return true;
}

/** Apply a persisted model (native column and row, EMU offsets). @throws {ExcelArgumentError} for a malformed one. */
export function setModel(a: AnchorData, value: AnchorModel): void {
  checkAnchor(a);
  checkAnchorEditable(a);
  checkModel(argumentChecks("Anchor.setModel"), "value", value);
  anchorSetModel(a, value);
}

export {
  anchorCol as getCol,
  anchorRow as getRow,
  anchorColWidth as getColWidth,
  anchorRowHeight as getRowHeight
};

/** A copy of the anchor as its model — the shape {@link setModel} takes: editing it changes nothing. */
export function getModel(a: AnchorData): AnchorModel {
  return anchorModel(a);
}

/** Options for {@link create}. */
export interface CreateAnchorOptions {
  /**
   * Added to both the column and the row of a string (`"B2"`) or `{ col, row }` address — `-1` turns a
   * one-based cell into the zero-based native position. Ignored for an anchor model. Default `0`.
   */
  offset?: number;
}

/**
 * Create an anchor on `worksheet` at `address` — a string (`"A1"`, or `"a1"`), an anchor model, or a
 * `{ col, row }` simple address. With no address the anchor sits at the top-left corner.
 *
 * @throws {ExcelInvalidAddressError} when a string is not a cell of the sheet (a sheet name is refused).
 * @throws {ExcelArgumentError} when a model or `{ col, row }` is off the grid, or `offset` is not a whole number.
 */
export function create(
  worksheet?: Worksheet,
  address?: AnchorInput | null,
  options?: CreateAnchorOptions
): AnchorData;
/**
 * @deprecated Pass the offset in the options object: `create(ws, address, { offset })`. Removed in 1.0.
 */
export function create(
  worksheet: Worksheet | undefined,
  address: AnchorInput | null | undefined,
  offset: number
): AnchorData;
export function create(
  worksheet?: Worksheet,
  address?: AnchorInput | null,
  offsetOrOptions?: number | CreateAnchorOptions
): AnchorData {
  // Optional: an anchor with no sheet measures columns and rows at their default size.
  if (worksheet !== undefined) {
    checkWorksheet(worksheet);
  }
  // Anything but an options object is the positional offset, handed on as the positional form
  // always did — including a `null` placeholder or a numeric string.
  const offset = isRecordArgument(offsetOrOptions) ? offsetOrOptions.offset : offsetOrOptions;
  const check = argumentChecks("Anchor.create");
  if (isRecordArgument(offsetOrOptions)) {
    check.number("options.offset", offset, { integer: true });
  }
  // What the anchor would be made of is checked before it is: an address that is not a cell, or a sheet name glued
  // onto the column letters (`Other!D5` read as column `OTHERD`), used to make an anchor of `null`s.
  if (typeof address === "string") {
    check.reference("address", address);
  } else if (address !== undefined && address !== null) {
    if (check.record("address", address) && "nativeCol" in address) {
      checkModel(check, "address", address);
    } else if (isRecordArgument(address)) {
      const simple = address as { col?: unknown; row?: unknown };
      const shift = typeof offset === "number" && Number.isFinite(offset) ? offset : 0;
      check.required("address.col", simple.col);
      check.number("address.col", simple.col, { min: -shift, max: MAX_COL - shift });
      check.required("address.row", simple.row);
      check.number("address.row", simple.row, { min: -shift, max: MAX_ROW - shift });
    }
  }
  return anchorCreate(worksheet, address, offset as number | undefined);
}

/** An anchor handle. */
export type { AnchorData as Handle } from "@excel/core/anchor";

// --- deprecated names ---

/** @deprecated Use {@link getCol}. Removed in 1.0. */
export const col = anchorCol;
/** @deprecated Use {@link getRow}. Removed in 1.0. */
export const row = anchorRow;
/** @deprecated Use {@link getColWidth}. Removed in 1.0. */
export const colWidth = anchorColWidth;
/** @deprecated Use {@link getRowHeight}. Removed in 1.0. */
export const rowHeight = anchorRowHeight;
/** @deprecated Use {@link getModel}. Removed in 1.0. */
export const model = getModel;
