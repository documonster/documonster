import { Cell, Column, ExcelColumnOutOfBoundsError, Row, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * `Column.getValues` (both bases) / `setValues` — reading and writing down a
 * column by key, letter or number.
 *
 * Two things here are easy to get wrong and invisible once wrong. The arrays are
 * *sparse*, so an empty cell is a hole rather than a `null`, and `toEqual`
 * cannot tell the two apart — the index assertions below use `Object.keys` for
 * that reason. And the reads must not materialise anything: every other member
 * of the namespace resolves its reference through `getColumn`, which pads
 * `ws._columns` up to the column asked for, so reading column `Z` of a
 * one-column sheet would silently declare 26 columns.
 */
function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", "a");
  Cell.setValue(ws, "A2", "b");
  Cell.setValue(ws, "A4", "d");
  return { wb, ws };
}

describe("Column.getValues with base 1", () => {
  it("is indexed by row number, with an empty leading slot", () => {
    const { ws } = sheet();

    const v = Column.getValues(ws, "A", { base: 1 });
    expect(v[1]).toBe("a");
    expect(v[2]).toBe("b");
    expect(v[4]).toBe("d");
    expect(Object.keys(v)).toEqual(["1", "2", "4"]);
  });

  it("resolves a key, a letter and a number to the same column", () => {
    const { ws } = sheet();
    Column.setDefinitions(ws, [{ key: "first", header: "First" }]);

    expect(Column.getValues(ws, "first", { base: 1 })).toEqual(
      Column.getValues(ws, "A", { base: 1 })
    );
    expect(Column.getValues(ws, 1, { base: 1 })).toEqual(Column.getValues(ws, "A", { base: 1 }));
  });

  it("composes with the Cell namespace, because the index is the row number", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "C1", 120);
    Cell.setValue(ws, "C2", -40);
    Cell.setValue(ws, "C3", 75);

    const colNumber = Column.getNumber(ws, "C");
    const reddened: number[] = [];
    Column.getValues(ws, "C", { base: 1 }).forEach((value, rowNumber) => {
      if (typeof value === "number" && value < 0) {
        Cell.setFont(ws, rowNumber, colNumber, { color: { argb: "FFC00000" } });
        reddened.push(rowNumber);
      }
    });

    expect(reddened).toEqual([2]);
    expect(Cell.getFont(ws, "C2")).toEqual({ color: { argb: "FFC00000" } });
  });

  it("is empty for a column that holds nothing", () => {
    const { ws } = sheet();
    expect(Column.getValues(ws, "Z", { base: 1 })).toEqual([]);
  });
});

describe("Column.getValues", () => {
  it("is 0-based: the value in row 1 is at index 0", () => {
    const { ws } = sheet();

    const v = Column.getValues(ws, "A");
    expect(v[0]).toBe("a");
    expect(v[1]).toBe("b");
    expect(v[3]).toBe("d");
    expect(Object.keys(v)).toEqual(["0", "1", "3"]);
  });

  it("is the 1-based read shifted by one", () => {
    const { ws } = sheet();
    expect(Column.getValues(ws, "A")).toEqual(Column.getValues(ws, "A", { base: 1 }).slice(1));
  });
});

describe("Column.setValues", () => {
  it("writes a dense array from row 1", () => {
    const { ws } = sheet();

    Column.setValues(ws, "B", [2, 3, 5]);

    expect(Cell.getValue(ws, "B1")).toBe(2);
    expect(Cell.getValue(ws, "B2")).toBe(3);
    expect(Cell.getValue(ws, "B3")).toBe(5);
  });

  it("reads a leading hole as row-number indexing, so a read round-trips", () => {
    const { ws } = sheet();

    Column.setValues(ws, "B", Column.getValues(ws, "A", { base: 1 }));

    expect(Column.getValues(ws, "B", { base: 1 })).toEqual(Column.getValues(ws, "A", { base: 1 }));
    expect(Object.keys(Column.getValues(ws, "B", { base: 1 }))).toEqual(["1", "2", "4"]);
  });

  it("treats a dense array and its row-number form as the same thing", () => {
    const { ws } = sheet();

    Column.setValues(ws, "B", [1, 2, 3]);
    Column.setValues(ws, "C", [, 1, 2, 3]);

    expect(Column.getValues(ws, "C")).toEqual(Column.getValues(ws, "B"));
  });

  it("leaves the rest of the column alone rather than replacing it", () => {
    const { ws } = sheet();

    // Row 4 is outside the array, and `Row.setValues` would have dropped it.
    Column.setValues(ws, "A", [1, 2]);

    expect(Column.getValues(ws, "A", { base: 1 })[4]).toBe("d");
  });

  it("keeps a value the source column has no value for, so it is not a copy", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "B3", "stale");

    // Column A holds rows 1, 2 and 4 — nothing at row 3, so nothing clears B3.
    Column.setValues(ws, "B", Column.getValues(ws, "A", { base: 1 }));

    expect(Column.getValues(ws, "B", { base: 1 })).toEqual([, "a", "b", "stale", "d"]);
  });

  it("clears a cell when the array carries null at its index", () => {
    const { ws } = sheet();

    Column.setValues(ws, "A", [, null]);

    expect(Cell.getValue(ws, "A1")).toBeNull();
    expect(Column.getValues(ws, "A", { base: 1 })[2]).toBe("b");
  });

  it("takes the wider Cell.setValue input, not just what a read returns", () => {
    const { ws } = sheet();

    Column.setValues(ws, "B", [
      { formula: "SUM(A1:A4)" },
      { richText: [{ text: "rich" }] },
      { error: "#N/A" },
      new Date(Date.UTC(2020, 0, 2))
    ]);

    expect(Cell.getValue(ws, "B1")).toMatchObject({ formula: "SUM(A1:A4)" });
    expect(Cell.getValue(ws, "B2")).toMatchObject({ richText: [{ text: "rich" }] });
    expect(Cell.getValue(ws, "B3")).toEqual({ error: "#N/A" });
    expect(Cell.getValue(ws, "B4")).toEqual(new Date(Date.UTC(2020, 0, 2)));
  });

  it("accepts a key as the reference", () => {
    const { ws } = sheet();
    Column.setDefinitions(ws, [{ key: "first" }, { key: "second" }]);

    Column.setValues(ws, "second", [7, 8]);

    expect(Cell.getValue(ws, "B1")).toBe(7);
    expect(Cell.getValue(ws, "B2")).toBe(8);
  });

  it("declares the column it writes to", () => {
    const { ws } = sheet();

    Column.setValues(ws, "D", [1]);

    expect(Column.getDefinitions(ws).length).toBe(4);
    expect(Cell.getValue(ws, "D1")).toBe(1);
  });
});

describe("row-number indexing is the round-trip form", () => {
  it("writes a base-1 read back unchanged even when row 1 is empty", () => {
    const { wb, ws } = sheet();
    const source = Workbook.addWorksheet(wb, "T");
    Cell.setValue(source, "A2", "x"); // nothing in row 1

    Column.setValues(ws, "B", Column.getValues(source, "A", { base: 1 }));

    expect(Cell.getValue(ws, "B1")).toBeNull();
    expect(Cell.getValue(ws, "B2")).toBe("x");
  });

  it("shifts Column.getValues up by a row when row 1 is empty, as Row does", () => {
    const { wb, ws } = sheet();
    const source = Workbook.addWorksheet(wb, "T");
    Cell.setValue(source, "A2", "x");

    // A 0-based array whose first slot is a hole is indistinguishable from a
    // row-number-indexed one, so this lands a row high. `Row` behaves the same,
    // which is why the docs point at `values` for writing a read back.
    Column.setValues(ws, "B", Column.getValues(source, "A"));
    Row.setValues(ws, 9, Row.getValues(source, 2));

    expect(Cell.getValue(ws, "B1")).toBe("x");
    expect(Cell.getValue(ws, "A9")).toBe("x");
  });
});

describe("column references are validated", () => {
  it("rejects a number outside Excel's column range", () => {
    const { ws } = sheet();

    for (const ref of [0, -1, 16385]) {
      expect(() => Column.getValues(ws, ref, { base: 1 })).toThrow(ExcelColumnOutOfBoundsError);
      expect(() => Column.getValues(ws, ref)).toThrow(ExcelColumnOutOfBoundsError);
      expect(() => Column.setValues(ws, ref, [1])).toThrow(ExcelColumnOutOfBoundsError);
    }
  });

  it("rejects a number that is not a whole column", () => {
    const { ws } = sheet();

    // Each of these used to index the cell array with something no row can hold,
    // so the read reported an empty column instead of a bad argument. `Infinity`
    // is the reason this matters: `getColumn` pads `_columns` in a `while` loop.
    for (const ref of [1.5, NaN, Infinity, -Infinity]) {
      expect(() => Column.getValues(ws, ref, { base: 1 })).toThrow(ExcelColumnOutOfBoundsError);
      expect(() => Column.setValues(ws, ref, [])).toThrow(ExcelColumnOutOfBoundsError);
    }
  });

  it("rejects a letter beyond XFD, the way it always did", () => {
    const { ws } = sheet();
    expect(() => Column.getValues(ws, "XFE", { base: 1 })).toThrow(ExcelColumnOutOfBoundsError);
  });

  it("does not mistake an Object.prototype member for a column key", () => {
    const { ws } = sheet();

    // `ws._keys` is a plain object, so a bare lookup finds these on the
    // prototype and reports a column number of `undefined`.
    for (const ref of ["toString", "constructor", "hasOwnProperty"]) {
      expect(() => Column.getValues(ws, ref, { base: 1 })).toThrow(ExcelColumnOutOfBoundsError);
    }
  });

  it("still resolves a declared key that shadows an Object.prototype member", () => {
    const { ws } = sheet();
    Column.setDefinitions(ws, [{ key: "toString" }]);

    expect(Column.getValues(ws, "toString", { base: 1 })).toEqual(
      Column.getValues(ws, "A", { base: 1 })
    );
  });
});

// That the two readers materialise nothing is asserted in
// `core/__tests__/read-side-effects.test.ts`, which fingerprints the saved bytes
// as well as the counters — a strictly stronger check than anything expressible
// here, and the file that owns that invariant for every reader in the library.

describe("values with an explicit base are inverse pairs", () => {
  it("Column: base 0 and base 1 read back what they wrote, even with row 1 empty", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "E");
    Column.setValues(ws, "A", [, "x", "y"], { base: 0 });
    expect(Cell.getValue(ws, "A1")).toBeNull();
    expect(Column.getValues(ws, "A")).toEqual([, "x", "y"]);
    Column.setValues(ws, "B", Column.getValues(ws, "A"), { base: 0 });
    expect(Column.getValues(ws, "B", { base: 1 })).toEqual(Column.getValues(ws, "A", { base: 1 }));
    // base 1 ignores index 0 even when present
    Column.setValues(ws, "C", ["ignored", "r1"], { base: 1 });
    expect(Column.getValues(ws, "C")).toEqual(["r1"]);
  });

  it("Row: the same two pairs, by column", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "E");
    Row.setValues(ws, 1, [, "b", "c"], { base: 0 });
    expect(Row.getValues(ws, 1)).toEqual([, "b", "c"]);
    Row.setValues(ws, 2, Row.getValues(ws, 1), { base: 0 });
    expect(Row.getValues(ws, 2, { base: 1 })).toEqual([, , "b", "c"]);
    Row.setValues(ws, 3, ["ignored", "a"], { base: 1 });
    expect(Row.getValues(ws, 3)).toEqual(["a"]);
  });
});

describe("Row and Column iteration are mirrors", () => {
  it("each visits the lines holding a value, or all up to the count", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "E");
    Cell.setValue(ws, "B2", 1);
    Cell.setValue(ws, "D4", 2);
    const cols: number[] = [];
    Column.each(ws, c => cols.push(c));
    expect(cols).toEqual([2, 4]);
    const all: number[] = [];
    Column.each(ws, c => all.push(c), { includeEmpty: true });
    expect(all).toEqual([1, 2, 3, 4]);
    const rows: number[] = [];
    Row.each(ws, r => rows.push(r));
    expect(rows).toEqual([2, 4]);
  });

  it("eachAddress with includeEmpty stops at the line's last cell in both directions", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "E");
    Cell.setValue(ws, "B2", 1);
    Cell.setValue(ws, "A5", 2); // row 5 exists, but has no cell in column B
    const down: string[] = [];
    Column.eachAddress(ws, "B", a => down.push(a), { includeEmpty: true });
    expect(down).toEqual(["B1", "B2"]);
    const across: string[] = [];
    Row.eachAddress(ws, 2, a => across.push(a), { includeEmpty: true });
    expect(across).toEqual(["A2", "B2"]);
  });
});
