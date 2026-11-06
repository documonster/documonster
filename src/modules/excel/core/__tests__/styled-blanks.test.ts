import { StyledBlankGrid, StyledBlankRuns, takeStyledBlanks } from "@excel/core/styled-blanks";
import type { StyledBlankArea } from "@excel/core/styled-blanks";
import { describe, expect, it } from "vitest";

/** Every cell a set of rectangles covers, as `row:column`, failing on a cell covered twice. */
function coverage(
  ranges: readonly (StyledBlankArea & { readonly styleIndex: number })[]
): Map<string, number> {
  const cells = new Map<string, number>();
  for (const range of ranges) {
    for (let row = range.firstRow; row <= range.lastRow; row++) {
      for (let column = range.firstColumn; column <= range.lastColumn; column++) {
        const key = `${row}:${column}`;
        expect(cells.has(key), `${key} covered twice`).toBe(false);
        cells.set(key, range.styleIndex);
      }
    }
  }
  return cells;
}

/** A small deterministic generator, so a failure reproduces. */
function random(seed: number): () => number {
  let state = seed;
  return () => {
    state = (state * 1103515245 + 12345) % 2147483648;
    return state / 2147483648;
  };
}

describe("takeStyledBlanks", () => {
  it("leaves the rectangles alone when no cell falls inside one", () => {
    const ranges = [{ firstRow: 0, lastRow: 3, firstColumn: 0, lastColumn: 3, styleIndex: 1 }];
    const result = takeStyledBlanks(ranges, [{ row: 5, column: 1 }]);
    expect(result.ranges).toEqual(ranges);
    expect(result.taken).toEqual([]);
  });

  it("cuts a cell out of the middle of a rectangle into four pieces", () => {
    const range = { firstRow: 0, lastRow: 2, firstColumn: 0, lastColumn: 2, styleIndex: 7 };
    const result = takeStyledBlanks([range], [{ row: 1, column: 1 }]);
    expect(result.taken).toEqual([{ cell: { row: 1, column: 1 }, range }]);
    expect(result.ranges).toEqual([
      { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 2, styleIndex: 7 },
      { firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 0, styleIndex: 7 },
      { firstRow: 1, lastRow: 1, firstColumn: 2, lastColumn: 2, styleIndex: 7 },
      { firstRow: 2, lastRow: 2, firstColumn: 0, lastColumn: 2, styleIndex: 7 }
    ]);
  });

  it("takes a cell named twice once", () => {
    const range = { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 1, styleIndex: 1 };
    const result = takeStyledBlanks(
      [range],
      [
        { row: 0, column: 0 },
        { row: 0, column: 0 }
      ]
    );
    expect(result.taken).toHaveLength(1);
    expect(result.ranges).toEqual([{ ...range, firstColumn: 1 }]);
  });

  it("removes exactly the cells it takes, and keeps every other cell's style", () => {
    // The property the writers rely on: the rectangles still describe every blank they described, minus the cells
    // handed back, and no cell twice.
    for (let seed = 1; seed <= 60; seed++) {
      const next = random(seed);
      const runs = new StyledBlankRuns();
      for (let row = 0; row < 12; row++) {
        for (let column = 0; column < 9; column++) {
          if (next() < 0.8) {
            runs.add(row, column, next() < 0.7 ? 1 : 2);
          }
        }
      }
      const ranges = runs.ranges();
      const before = coverage(ranges);
      const cells = Array.from({ length: 1 + Math.floor(next() * 20) }, () => ({
        row: Math.floor(next() * 14),
        column: Math.floor(next() * 11)
      }));
      const { ranges: remaining, taken } = takeStyledBlanks(ranges, cells);
      const after = coverage(remaining);

      const expected = new Map(before);
      const takenKeys = new Set<string>();
      for (const { cell, range } of taken) {
        const key = `${cell.row}:${cell.column}`;
        expect(takenKeys.has(key), `seed ${seed}: ${key} taken twice`).toBe(false);
        takenKeys.add(key);
        expect(before.get(key), `seed ${seed}: ${key}`).toBe(range.styleIndex);
        expected.delete(key);
      }
      for (const cell of cells) {
        const key = `${cell.row}:${cell.column}`;
        expect(takenKeys.has(key), `seed ${seed}: ${key}`).toBe(before.has(key));
      }
      expect(after, `seed ${seed}`).toEqual(expected);
    }
  });
});

describe("StyledBlankGrid", () => {
  const STYLES = [{ numFmt: "a" }, { numFmt: "b" }, { numFmt: "c" }];

  /** The cells a grid covers, as `row:column` → the style's `numFmt`, failing on a cell covered twice. */
  function cellsOf(grid: StyledBlankGrid): Map<string, string | undefined> {
    const cells = new Map<string, string | undefined>();
    for (const range of grid.ranges()) {
      for (let row = range.firstRow; row <= range.lastRow; row++) {
        for (let column = range.firstColumn; column <= range.lastColumn; column++) {
          const key = `${row}:${column}`;
          expect(cells.has(key), `${key} covered twice`).toBe(false);
          cells.set(key, range.style.numFmt as string | undefined);
        }
      }
    }
    return cells;
  }

  /** The same edit on a plain map of cells, as the cells themselves would move. */
  function spliceCells(
    cells: Map<string, string | undefined>,
    axis: "row" | "column",
    start: number,
    deleted: number,
    inserted: number,
    lastLine: number
  ): Map<string, string | undefined> {
    const out = new Map<string, string | undefined>();
    for (const [key, style] of cells) {
      const [row, column] = key.split(":").map(Number);
      const line = axis === "row" ? row : column;
      if (line >= start && line < start + deleted) {
        continue;
      }
      const moved = line >= start + deleted ? line + inserted - deleted : line;
      if (moved > lastLine) {
        continue;
      }
      out.set(axis === "row" ? `${moved}:${column}` : `${row}:${moved}`, style);
    }
    return out;
  }

  it("answers, takes rows and splices exactly as the cells would", () => {
    for (let seed = 1; seed <= 40; seed++) {
      const next = random(seed);
      const runs = new StyledBlankRuns();
      for (let row = 0; row < 14; row++) {
        for (let column = 0; column < 10; column++) {
          if (next() < 0.75) {
            runs.add(row, column, next() < 0.6 ? 0 : next() < 0.5 ? 1 : 2);
          }
        }
      }
      const grid = StyledBlankGrid.from(
        runs.ranges().map(range => ({ ...range, style: STYLES[range.styleIndex] }))
      );
      let expected = cellsOf(grid);
      for (let step = 0; step < 12; step++) {
        const choice = next();
        if (choice < 0.5) {
          const row = Math.floor(next() * 16);
          const taken = grid.takeRow(row);
          for (const run of taken) {
            for (let column = run.firstColumn; column <= run.lastColumn; column++) {
              expect(expected.get(`${row}:${column}`), `seed ${seed}`).toBe(run.style.numFmt);
              expected.delete(`${row}:${column}`);
            }
          }
        } else {
          const axis = choice < 0.75 ? "row" : "column";
          const start = Math.floor(next() * 14);
          const deleted = Math.floor(next() * 4);
          const inserted = Math.floor(next() * 4);
          grid.splice(axis, start, deleted, inserted, 20);
          expected = spliceCells(expected, axis, start, deleted, inserted, 20);
        }
        expect(cellsOf(grid), `seed ${seed}, step ${step}`).toEqual(expected);
        for (let row = 0; row < 22; row++) {
          for (let column = 0; column < 22; column++) {
            expect(grid.at(row, column)?.style.numFmt, `seed ${seed} ${row}:${column}`).toBe(
              expected.get(`${row}:${column}`)
            );
          }
        }
      }
    }
  });

  it("answers between takes without being rebuilt, and agrees once it is", () => {
    for (let seed = 1; seed <= 20; seed++) {
      const next = random(seed);
      const grid = StyledBlankGrid.from([
        { firstRow: 0, lastRow: 29, firstColumn: 0, lastColumn: 2, style: STYLES[0] },
        { firstRow: 10, lastRow: 39, firstColumn: 4, lastColumn: 4, style: STYLES[1] }
      ]);
      const live = new Set(Array.from({ length: 40 }, (_, row) => row));
      for (let step = 0; step < 45; step++) {
        const row = Math.floor(next() * 42);
        expect(grid.takeRow(row).length > 0, `seed ${seed} row ${row}`).toBe(live.delete(row));
        expect(grid.runsAt(row)).toEqual([]);
        expect(grid.isEmpty()).toBe(live.size === 0);
      }
      for (let row = 0; row < 42; row++) {
        expect(grid.at(row, 0) !== undefined).toBe(live.has(row) && row < 30);
        expect(grid.at(row, 4) !== undefined).toBe(live.has(row) && row >= 10);
      }
      const lastRow = Math.max(-1, ...live);
      expect(grid.extent()?.lastRow).toBe(lastRow === -1 ? undefined : lastRow);
      const rows = new Set(
        grid
          .ranges()
          .flatMap(r =>
            Array.from({ length: r.lastRow - r.firstRow + 1 }, (_, i) => r.firstRow + i)
          )
      );
      expect(rows).toEqual(live);
    }
  });

  it("takes every row of a tall band in any order, and agrees with the cells at each half", () => {
    // Each take only records the row; the bands are cut once, when the grid is asked as a whole. Splitting the band per
    // row instead spliced the band array on every take.
    const rows = 20_000;
    const grid = StyledBlankGrid.from([
      { firstRow: 0, lastRow: rows - 1, firstColumn: 0, lastColumn: 3, style: STYLES[0] }
    ]);
    const order = Array.from({ length: rows }, (_, row) => row);
    const next = random(7);
    for (let i = order.length - 1; i > 0; i--) {
      const j = Math.floor(next() * (i + 1));
      [order[i], order[j]] = [order[j], order[i]];
    }
    const half = new Set(order.slice(0, rows / 2));
    for (const row of half) {
      expect(grid.takeRow(row)).toHaveLength(1);
    }
    let covered = 0;
    for (const range of grid.ranges()) {
      for (let row = range.firstRow; row <= range.lastRow; row++) {
        expect(half.has(row)).toBe(false);
        covered++;
      }
    }
    expect(covered).toBe(rows / 2);
    for (const row of order.slice(rows / 2)) {
      expect(grid.takeRow(row)).toHaveLength(1);
    }
    expect(grid.isEmpty()).toBe(true);
    expect(grid.ranges()).toEqual([]);
    expect(grid.extent()).toBeUndefined();
  });
});
