/**
 * Which table a formula sits in, for writing an unnamed this-row reference (`[@x]`) with that table's name.
 *
 * Shared by both writers so the XLSX and XLSB outputs agree on the enclosing table — see `toFileFormula`'s `table`
 * option for why the name has to be written.
 */
import { decodeCell, decodeRange } from "@excel/utils/address";

/** A table as far as locating it needs: its name and the range it occupies. */
export interface TableExtent {
  readonly name?: unknown;
  /** The range the table occupies once header and totals rows are counted, where the model has it. */
  readonly tableRef?: unknown;
  /** The caller's anchor or range. */
  readonly ref?: unknown;
}

/**
 * A lookup from a cell address to the name of the table whose range contains it, or `undefined`.
 *
 * A table whose range does not decode is skipped rather than guessed at.
 */
export function tableNameAt(
  tables: readonly TableExtent[] | undefined
): (address: string) => string | undefined {
  const extents: { name: string; top: number; left: number; bottom: number; right: number }[] = [];
  for (const table of tables ?? []) {
    const ref = typeof table.tableRef === "string" ? table.tableRef : table.ref;
    if (typeof table.name !== "string" || typeof ref !== "string") {
      continue;
    }
    try {
      const range = decodeRange(ref);
      extents.push({
        name: table.name,
        top: range.s.r,
        left: range.s.c,
        bottom: range.e.r,
        right: range.e.c
      });
    } catch {
      // Not a range: no cell is inside it.
    }
  }
  return address => {
    if (extents.length === 0) {
      return undefined;
    }
    let cell: { r: number; c: number };
    try {
      cell = decodeCell(address);
    } catch {
      return undefined;
    }
    return extents.find(
      extent =>
        cell.r >= extent.top &&
        cell.r <= extent.bottom &&
        cell.c >= extent.left &&
        cell.c <= extent.right
    )?.name;
  };
}
