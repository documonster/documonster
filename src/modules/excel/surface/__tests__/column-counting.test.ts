/**
 * "Exists" means one thing to every `Column` counting member, as it does to `Row`'s: some row has a cell in the column,
 * or the column has a setting of its own. `Column.has` already said so; `getCount`, `getLast` and
 * `each({ includeEmpty })` looked at cells only, so a column with only a width was skipped by a loop up to
 * `getLast`.
 */
import { Cell, Column, Row, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

function sheet() {
  const ws = Workbook.addWorksheet(Workbook.create(), "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "D1", 2);
  Row.setHeight(ws, 5, 20);
  Column.setWidth(ws, "H", 12);
  return ws;
}

describe("Column counting agrees with Column.has, as Row's does with Row.has", () => {
  it("Row, for reference", () => {
    const ws = sheet();
    const each: number[] = [];
    Row.each(ws, r => each.push(r), { includeEmpty: true });
    expect([Row.has(ws, 5), Row.getLast(ws), Row.getCount(ws), each]).toEqual([
      true,
      5,
      5,
      [1, 2, 3, 4, 5]
    ]);
  });

  it("Column", () => {
    const ws = sheet();
    const each: number[] = [];
    Column.each(ws, c => each.push(c), { includeEmpty: true });
    expect([Column.has(ws, "H"), Column.getLast(ws), Column.getCount(ws), each]).toEqual([
      true,
      8,
      8,
      [1, 2, 3, 4, 5, 6, 7, 8]
    ]);
    expect(Column.getActualCount(ws)).toBe(2);
    const used: number[] = [];
    Column.each(ws, c => used.push(c));
    expect(used).toEqual([1, 4]);
  });

  it("a column record with nothing set does not count", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "S");
    Cell.setValue(ws, "B1", 1);
    Column.setWidth(ws, "F", 5);
    Column.setWidth(ws, "F", undefined);
    expect([Column.has(ws, "F"), Column.getLast(ws), Column.getCount(ws)]).toEqual([false, 2, 2]);
  });

  it("an empty sheet", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "S");
    expect([Column.getLast(ws), Column.getCount(ws)]).toEqual([undefined, 0]);
  });
});
