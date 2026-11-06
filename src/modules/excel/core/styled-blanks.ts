/**
 * Cells that carry formatting and no value, as rectangles.
 *
 * **One representation, both containers.** `<c r="A9" s="3"/>` in XLSX and `BrtCellBlank` in XLSB are the same thing —
 * a cell whose only content is its format — and Excel writes one per cell of a formatted region. A sheet with a fill
 * applied past its data therefore carries tens of thousands of them: measured at 62,400 cells and 36.8 MB retained for
 * an eight-column region against 200 rows of actual data, and 322,520 cells and 186 MB in the binary form.
 *
 * `blankCells: "collapse"` accumulates them here instead of giving each one a cell in the model, and **both writers
 * expand them again** — which is what makes the option lossless rather than a discard with a nicer name. That is also
 * why this lives in `core/` rather than beside either reader: an option whose fidelity depended on which container the
 * workbook came from would be the worse API, and sharing the representation is the only way to avoid it.
 *
 * The rectangles are exact. Every cell inside one was a styled blank with that style, and no cell outside one was, so
 * a writer reproduces the records rather than approximating them.
 */

import type { Style } from "@excel/types";

/** A rectangle of blank cells that share one style. */
export interface StyledBlankRange {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  /** The style, as the reader that collected it identifies one: an XLSB `ixfe` or an XLSX `styleId`. */
  readonly styleIndex: number;
}

/**
 * A styled blank rectangle as the **workbook model** carries it.
 *
 * The collector's {@link StyledBlankRange} holds a `styleIndex` — an `ixfe` or an XLSX `styleId` — which is meaningful
 * only against the styles table it was read from. By the time the range reaches `WorksheetModel` that index has been
 * resolved to the style itself, because a style table is rebuilt on write and an index into the old one would name the
 * wrong format.
 *
 * **Named because the public field used to be `readonly unknown[]`.** `Workbook.getModel` and `setModel` are public, so a
 * caller reading or building this had nothing to reference and both writers reached for their own private cast — which is
 * how the reader's shape and the model's shape came to differ in the first place.
 */
export interface StyledBlankRangeModel {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  /** The resolved style, not an index into a table that no longer exists. */
  readonly style: Partial<Style>;
}

/**
 * The extent a styled blank rectangle covers, whatever else it carries — the collector's index or the model's style.
 * Zero-based, inclusive.
 */
export interface StyledBlankArea {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

/** A cell, zero-based, as the rectangles address one. */
export interface StyledBlankCell {
  readonly row: number;
  readonly column: number;
}

/**
 * Take cells back out of the rectangles they were collapsed into.
 *
 * **A rectangle may only cover cells that hold nothing but their format**, and whether a cell does is not always known
 * when it is collapsed: a hyperlink or a comment reaches a cell only after the cells have been read, from a later part of
 * the sheet or from a part of its own. A cell that turns out to carry one is no longer a styled blank, so it is taken out
 * here and handed back to the reader, which gives it a cell — with its style, as `"keep"` would have — before attaching
 * the annotation. Left inside the rectangle, an XLSX hyperlink found no cell and was dropped, and every other annotation
 * was given a new cell without the format — which the writers kept, since they fill in only cells that do not exist.
 *
 * Returns the rectangles that remain, still exact, and each taken cell with the rectangle it came from. A cell outside
 * every rectangle is not taken, and one named twice is taken once. A rectangle with a cell removed is split into at most
 * four: the rows above the cell, the rows below it, and the two runs beside it in its own row.
 */
export function takeStyledBlanks<R extends StyledBlankArea, C extends StyledBlankCell>(
  ranges: readonly R[],
  cells: Iterable<C>
): { readonly ranges: R[]; readonly taken: { readonly cell: C; readonly range: R }[] } {
  // Sorted by row so each rectangle finds the cells in its rows by binary search: the cost is a search per rectangle
  // plus the cells actually taken, rather than every cell against every rectangle.
  const wanted = [...cells].sort(
    (left, right) => left.row - right.row || left.column - right.column
  );
  const remaining: R[] = [];
  const taken: { cell: C; range: R }[] = [];
  if (wanted.length === 0) {
    return { ranges: [...ranges], taken };
  }
  for (const range of ranges) {
    let low = 0;
    let high = wanted.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (wanted[middle].row < range.firstRow) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    const inside: C[] = [];
    for (let index = low; index < wanted.length && wanted[index].row <= range.lastRow; index++) {
      const cell = wanted[index];
      const previous = inside[inside.length - 1];
      if (
        cell.column >= range.firstColumn &&
        cell.column <= range.lastColumn &&
        (previous === undefined || previous.row !== cell.row || previous.column !== cell.column)
      ) {
        inside.push(cell);
      }
    }
    if (inside.length === 0) {
      remaining.push(range);
      continue;
    }
    for (const cell of inside) {
      taken.push({ cell, range });
    }
    carve(range, inside, remaining);
  }
  return { ranges: remaining, taken };
}

/** Split `range` around `cells` (all inside it, sorted, distinct) and push the pieces that are left. */
function carve<R extends StyledBlankArea>(
  range: R,
  cells: readonly StyledBlankCell[],
  out: R[]
): void {
  const piece = (firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): R =>
    ({ ...range, firstRow, lastRow, firstColumn, lastColumn }) as R;
  // Row by row: the rows between two cells' rows are whole, and a cell's own row is cut into the runs beside it.
  let nextRow = range.firstRow;
  let index = 0;
  while (index < cells.length) {
    const row = cells[index].row;
    if (row > nextRow) {
      out.push(piece(nextRow, row - 1, range.firstColumn, range.lastColumn));
    }
    let column = range.firstColumn;
    for (; index < cells.length && cells[index].row === row; index++) {
      if (cells[index].column > column) {
        out.push(piece(row, row, column, cells[index].column - 1));
      }
      column = cells[index].column + 1;
    }
    if (column <= range.lastColumn) {
      out.push(piece(row, row, column, range.lastColumn));
    }
    nextRow = row + 1;
  }
  if (nextRow <= range.lastRow) {
    out.push(piece(nextRow, range.lastRow, range.firstColumn, range.lastColumn));
  }
}

/**
 * Styled blank cells, accumulated into as few rectangles as describe them exactly.
 *
 * **Two passes, and both are needed for the shape real files have.** Cells arrive row by row in ascending column
 * order, so a horizontal run is recognised by looking only at the previous cell. Vertical merging then joins runs
 * in adjacent rows that cover the same columns with the same style, which is what turns a formatted column into
 * one rectangle instead of one run per row — the difference between a handful of objects and sixteen thousand.
 *
 * The result is exact: every cell in a rectangle was a `BrtCellBlank` with that style, and no cell outside one
 * was. That is what lets the writer reproduce the records rather than approximate them, and it is why this is a
 * different thing from dropping the cells and reporting how many were lost.
 */
export class StyledBlankRuns {
  /** Finished rectangles, and the runs of the row currently being read. */
  private readonly closed: StyledBlankRange[] = [];
  private open: StyledBlankRange[] = [];
  private openRow = -1;

  add(row: number, column: number, styleIndex: number): void {
    if (row !== this.openRow) {
      this.closeRow();
      this.openRow = row;
    }
    const last = this.open[this.open.length - 1];
    if (last !== undefined && last.styleIndex === styleIndex && last.lastColumn === column - 1) {
      this.open[this.open.length - 1] = { ...last, lastColumn: column };
      return;
    }
    this.open.push({
      firstRow: row,
      lastRow: row,
      firstColumn: column,
      lastColumn: column,
      styleIndex
    });
  }

  ranges(): readonly StyledBlankRange[] {
    this.closeRow();
    return this.closed;
  }

  /**
   * Merge the finished row's runs into the rectangles above them where they line up exactly.
   *
   * Only an exact match extends a rectangle — same first column, same last column, same style, and the row
   * directly below. A run that merely overlaps starts its own rectangle, because a rectangle that covered a cell
   * which was not blank would make the writer invent a record.
   */
  private closeRow(): void {
    for (const run of this.open) {
      const above = this.closed.find(
        candidate =>
          candidate.lastRow === run.firstRow - 1 &&
          candidate.firstColumn === run.firstColumn &&
          candidate.lastColumn === run.lastColumn &&
          candidate.styleIndex === run.styleIndex
      );
      if (above === undefined) {
        this.closed.push(run);
        continue;
      }
      this.closed[this.closed.indexOf(above)] = { ...above, lastRow: run.lastRow };
    }
    this.open = [];
  }
}
