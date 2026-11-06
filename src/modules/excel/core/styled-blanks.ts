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

/** A run of columns (zero-based, inclusive) in one band of a {@link StyledBlankGrid}, and the style its cells share. */
export interface StyledBlankRun {
  readonly firstColumn: number;
  readonly lastColumn: number;
  readonly style: Partial<Style>;
}

/** Rows `first`..`last` (zero-based, inclusive), all covering the same runs. Empty runs: the rows hold no blank. */
interface StyledBlankBand {
  readonly first: number;
  readonly last: number;
  readonly runs: readonly StyledBlankRun[];
}

/** Where lines `first`..`last` go in a splice of `deleted` lines at `start` replaced by `inserted`: none, one or two spans. */
function spliceLines(
  first: number,
  last: number,
  start: number,
  deleted: number,
  inserted: number,
  lastLine: number
): { readonly first: number; readonly last: number }[] {
  const shift = inserted - deleted;
  const out: { first: number; last: number }[] = [];
  if (first < start) {
    out.push({ first, last: Math.min(last, start - 1) });
  }
  const kept = Math.max(first, start + deleted);
  if (kept <= last && kept + shift <= lastLine) {
    out.push({ first: kept + shift, last: Math.min(last + shift, lastLine) });
  }
  return out;
}

/** Rows the bands cover with at least one run. */
function countRows(bands: readonly StyledBlankBand[]): number {
  let rows = 0;
  for (const band of bands) {
    if (band.runs.length > 0) {
      rows += band.last - band.first + 1;
    }
  }
  return rows;
}

/**
 * The styled blanks of a sheet that holds them as rectangles, indexed for the questions the sheet asks of them.
 *
 * **Held as bands of rows, not as the rectangles.** A sheet asks by row — which blanks are in this row, so it can be
 * given its cells; is this cell one — and a row is answered here by one binary search, where a list of rectangles has to
 * be walked whole. That difference is the whole cost of a collapsed sheet: a manually banded region is one rectangle per
 * row, and a walk per cell turned reading the format of every cell, or writing a value down a column, into
 * rectangles × cells. The rectangles are what a model and the writers carry, so they are rebuilt on the way out.
 *
 * Bands are disjoint and in row order, and never change while rows are taken out: a taken row is only recorded
 * (`taken`), and every question about a row asks that set first. Splitting the band instead cost a splice of the band
 * array per row, which is quadratic once rows are settled in no particular order — every other row of a tall region, or
 * the targets of a formula's writeback. The bands are cut at the taken rows (`compact`) only when the grid is asked as a
 * whole — for its rectangles, its extent, or to move — which happens once per write or structural edit.
 *
 * **Never larger than the cells it stands for.** A band holds one run for each rectangle crossing its rows, so a sheet
 * whose rectangles overlap in rows without lining up — styles side by side, starting on different rows — has more runs
 * than rectangles. But each run covers at least one cell of its band, so there are never more runs than the cells
 * `"keep"` would have held, and each is a fraction of a cell's size.
 */
export class StyledBlankGrid {
  private bands: StyledBlankBand[];
  private cachedRanges: StyledBlankRangeModel[] | undefined;
  /** Rows (zero-based) taken out of a band since the bands were last cut — see {@link compact}. */
  private readonly taken = new Set<number>();
  /** Rows that hold a blank and are not taken. */
  private liveRows: number;

  private constructor(bands: StyledBlankBand[]) {
    this.bands = bands;
    this.liveRows = countRows(bands);
  }

  /** The grid the rectangles describe. They are taken as exact: no cell in two of them. */
  static from(ranges: readonly StyledBlankRangeModel[]): StyledBlankGrid {
    const edges = [...new Set(ranges.flatMap(range => [range.firstRow, range.lastRow + 1]))].sort(
      (left, right) => left - right
    );
    const starting = [...ranges].sort((left, right) => left.firstRow - right.firstRow);
    const bands: StyledBlankBand[] = [];
    let active: StyledBlankRangeModel[] = [];
    let next = 0;
    for (let index = 0; index + 1 < edges.length; index++) {
      const first = edges[index];
      const last = edges[index + 1] - 1;
      active = active.filter(range => range.lastRow >= first);
      while (next < starting.length && starting[next].firstRow === first) {
        active.push(starting[next++]);
      }
      if (active.length > 0) {
        const runs = active
          .map(range => ({
            firstColumn: range.firstColumn,
            lastColumn: range.lastColumn,
            style: range.style
          }))
          .sort((left, right) => left.firstColumn - right.firstColumn);
        bands.push({ first, last, runs });
      }
    }
    const grid = new StyledBlankGrid(bands);
    grid.cachedRanges = [...ranges];
    return grid;
  }

  /**
   * The rectangles, rebuilt from the bands: runs that line up exactly across adjacent bands, with the same style, are
   * one rectangle again. Cached until the grid changes.
   */
  ranges(): readonly StyledBlankRangeModel[] {
    if (this.cachedRanges !== undefined) {
      return this.cachedRanges;
    }
    this.compact();
    type Open = {
      firstRow: number;
      lastRow: number;
      firstColumn: number;
      lastColumn: number;
      style: Partial<Style>;
    };
    const out: Open[] = [];
    // The rectangles still open below the previous band, in column order like the runs of every band — so the run that
    // continues one is found by walking both lists together, not by searching.
    let open: Open[] = [];
    let previousLast = Number.NaN;
    for (const band of this.bands) {
      if (band.runs.length === 0) {
        continue;
      }
      const above = band.first === previousLast + 1 ? open : [];
      const nextOpen: Open[] = [];
      let candidate = 0;
      for (const run of band.runs) {
        while (candidate < above.length && above[candidate].firstColumn < run.firstColumn) {
          candidate++;
        }
        const match = above[candidate];
        if (
          match !== undefined &&
          match.firstColumn === run.firstColumn &&
          match.lastColumn === run.lastColumn &&
          match.style === run.style
        ) {
          match.lastRow = band.last;
          nextOpen.push(match);
        } else {
          const rectangle: Open = {
            firstRow: band.first,
            lastRow: band.last,
            firstColumn: run.firstColumn,
            lastColumn: run.lastColumn,
            style: run.style
          };
          out.push(rectangle);
          nextOpen.push(rectangle);
        }
      }
      open = nextOpen;
      previousLast = band.last;
    }
    this.cachedRanges = out;
    return out;
  }

  isEmpty(): boolean {
    return this.liveRows === 0;
  }

  /** The last row and the last column (zero-based) any blank reaches, or `undefined` when there is none. */
  extent(): { readonly lastRow: number; readonly lastColumn: number } | undefined {
    this.compact();
    let lastRow = -1;
    let lastColumn = -1;
    for (const band of this.bands) {
      if (band.runs.length > 0) {
        lastRow = band.last;
        lastColumn = Math.max(lastColumn, band.runs[band.runs.length - 1].lastColumn);
      }
    }
    return lastRow === -1 ? undefined : { lastRow, lastColumn };
  }

  /** The runs in a row (zero-based); empty when it holds no blank. */
  runsAt(row: number): readonly StyledBlankRun[] {
    if (this.taken.has(row)) {
      return [];
    }
    const index = this.bandIndex(row);
    return index === -1 ? [] : this.bands[index].runs;
  }

  /** The run covering a cell (zero-based), or `undefined`. */
  at(row: number, column: number): StyledBlankRun | undefined {
    const runs = this.runsAt(row);
    let low = 0;
    let high = runs.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (runs[middle].lastColumn < column) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    const run = runs[low];
    return run !== undefined && run.firstColumn <= column ? run : undefined;
  }

  /** Take a row (zero-based) out of the grid, returning the runs it held. */
  takeRow(row: number): readonly StyledBlankRun[] {
    const runs = this.runsAt(row);
    if (runs.length === 0) {
      return runs;
    }
    this.taken.add(row);
    this.liveRows--;
    this.cachedRanges = undefined;
    return runs;
  }

  /**
   * Move the blanks as a row or column splice moves the cells: zero-based along `axis`, the `deleted` lines from `start`
   * go, `inserted` new ones take their place, and every line after them moves by the difference. A band or run across
   * the edit is cut rather than stretched, because the new lines are empty. Nothing past `lastLine`, the edge of the
   * grid, is kept: a sheet refuses a splice that would push a blank there before it starts (`checkSpliceFits`), so this
   * only keeps the grid total.
   */
  splice(
    axis: "row" | "column",
    start: number,
    deleted: number,
    inserted: number,
    lastLine: number
  ): void {
    this.compact();
    const bands: StyledBlankBand[] = [];
    for (const band of this.bands) {
      if (band.runs.length === 0) {
        continue;
      }
      if (axis === "row") {
        for (const lines of spliceLines(
          band.first,
          band.last,
          start,
          deleted,
          inserted,
          lastLine
        )) {
          bands.push({ first: lines.first, last: lines.last, runs: band.runs });
        }
        continue;
      }
      const runs = band.runs.flatMap(run =>
        spliceLines(run.firstColumn, run.lastColumn, start, deleted, inserted, lastLine).map(
          lines => ({
            firstColumn: lines.first,
            lastColumn: lines.last,
            style: run.style
          })
        )
      );
      if (runs.length > 0) {
        bands.push({ first: band.first, last: band.last, runs });
      }
    }
    this.bands = bands;
    this.liveRows = countRows(bands);
    this.cachedRanges = undefined;
  }

  /** Cut the bands at the taken rows, so they describe the grid by themselves again. */
  private compact(): void {
    if (this.taken.size === 0) {
      return;
    }
    const rows = [...this.taken].sort((left, right) => left - right);
    const bands: StyledBlankBand[] = [];
    let next = 0;
    for (const band of this.bands) {
      if (band.runs.length === 0) {
        continue;
      }
      while (next < rows.length && rows[next] < band.first) {
        next++;
      }
      let first = band.first;
      for (; next < rows.length && rows[next] <= band.last; next++) {
        if (rows[next] > first) {
          bands.push({ first, last: rows[next] - 1, runs: band.runs });
        }
        first = rows[next] + 1;
      }
      if (first <= band.last) {
        bands.push({ first, last: band.last, runs: band.runs });
      }
    }
    this.bands = bands;
    this.taken.clear();
  }

  /**
   * Visit each blank cell (zero-based) inside `area`, row by row. Costs the bands the area's rows cross and the cells
   * visited, not the area: a range over the whole sheet with two blanks in it visits two cells.
   */
  forEachCellIn(area: StyledBlankArea, visit: (row: number, column: number) => void): void {
    for (let index = this.firstBandFrom(area.firstRow); index < this.bands.length; index++) {
      const band = this.bands[index];
      if (band.first > area.lastRow) {
        break;
      }
      const top = Math.max(band.first, area.firstRow);
      const bottom = Math.min(band.last, area.lastRow);
      for (const run of band.runs) {
        const left = Math.max(run.firstColumn, area.firstColumn);
        const right = Math.min(run.lastColumn, area.lastColumn);
        for (let row = top; row <= bottom && left <= right; row++) {
          if (this.taken.has(row)) {
            continue;
          }
          for (let column = left; column <= right; column++) {
            visit(row, column);
          }
        }
      }
    }
  }

  /** The first band (by index) whose last row is at or after `row`. */
  private firstBandFrom(row: number): number {
    let low = 0;
    let high = this.bands.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (this.bands[middle].last < row) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    return low;
  }

  /** The band holding a row (zero-based), or -1. */
  private bandIndex(row: number): number {
    const low = this.firstBandFrom(row);
    return low < this.bands.length && this.bands[low].first <= row ? low : -1;
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
