import { ExcelError, InvalidAddressError } from "@excel/errors";
import { colCache } from "@excel/utils/col-cache";

/**
 * How a worksheet's `<row r>` and `<c r>` are read — one set of rules for the full reader (`RowXform`) and the
 * streaming one, so the two cannot place a cell differently.
 *
 * Both attributes are optional (ECMA-376 §18.3.1.73, §18.3.1.4): an absent one means the row after the previous row,
 * or the cell after the previous cell of its row. A present one that is not a valid reference to its own row is an
 * error naming the value, rather than a guess at where the cell belongs.
 */

/** Excel's last row and column. */
const MAX_ROW = 1048576;
const MAX_COL = 16384;

/** A reference spelled with `$` or leading zeros — `decodeAddress` reads these; anything else it would half-read. */
const LOOSE_REFERENCE_RX = /^\$?[A-Z]{1,3}\$?\d+$/;

/** The row a `<row r>` names, or the one after `previousRow` when it is absent or empty. */
export function readRowNumber(r: string | undefined, previousRow: number): number {
  let row = 0;
  if (!r) {
    row = previousRow + 1;
  } else {
    // `xsd:unsignedInt`: surrounding whitespace, an optional `+`, then digits — leading zeros allowed.
    const digits = r.trim().replace(/^\+/, "");
    for (let i = 0; i < digits.length; i++) {
      const c = digits.charCodeAt(i);
      if (c < 48 || c > 57) {
        row = 0;
        break;
      }
      row = row * 10 + c - 48;
    }
  }
  if (row < 1 || row > MAX_ROW) {
    throw new ExcelError(`Invalid row number "${r ?? row}": expected 1 to ${MAX_ROW}`);
  }
  return row;
}

/**
 * The 1-based column a `<c r>` names, or the one after `previousCol` when it is absent or empty.
 *
 * A present reference must be a whole cell reference and must lie in `row`, the row it is written in.
 */
export function readCellColumn(r: string | undefined, row: number, previousCol: number): number {
  let col: number;
  if (!r) {
    col = previousCol + 1;
  } else {
    let refRow = colCache.decodePlainRow(r);
    if (refRow > 0) {
      col = colCache.decodeCol(r);
    } else if (LOOSE_REFERENCE_RX.test(r)) {
      const decoded = colCache.decodeAddress(r);
      col = decoded.col;
      refRow = decoded.row;
    } else {
      throw new InvalidAddressError(r);
    }
    if (refRow !== row) {
      throw new InvalidAddressError(r, `cell ${r} is written inside row ${row}`);
    }
  }
  if (col < 1 || col > MAX_COL) {
    throw new InvalidAddressError(r ?? String(col), `column ${col} is outside 1 to ${MAX_COL}`);
  }
  return col;
}
