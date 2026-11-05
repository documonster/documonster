/**
 * A pivot table's source cannot be removed from under it, two new pivot tables cannot share cells, and the model is
 * checked before anything is added.
 */
import {
  ExcelArgumentError,
  Pivot,
  ExcelPivotTableError,
  Table,
  Workbook,
  Worksheet
} from "@excel";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const data = Workbook.addWorksheet(wb, "Data");
  Worksheet.addAoa(data, [
    ["Region", "Sales"],
    ["N", 1],
    ["S", 2]
  ]);
  const out = Workbook.addWorksheet(wb, "Pivot");
  return { wb, data, out };
}

describe("a pivot table's source", () => {
  it("its sheet cannot be removed while the pivot table reads it", () => {
    const { wb, data, out } = book();
    const pivot = Pivot.add(out, { sourceSheet: data, rows: ["Region"], values: ["Sales"] });
    expect(() => Workbook.removeWorksheet(wb, data)).toThrow(ExcelPivotTableError);
    expect(Workbook.getWorksheet(wb, "Data")).toBe(data);
    Pivot.remove(out, pivot);
    expect(Workbook.removeWorksheet(wb, data)).toBe(true);
  });

  it("its table cannot be removed, nor the sheet that holds the table", () => {
    const { wb, data, out } = book();
    const table = Table.add(data, {
      name: "Src",
      ref: "D1",
      columns: [{ name: "k" }, { name: "v" }],
      rows: [["a", 1]]
    });
    Pivot.add(out, { sourceTable: table, rows: ["k"], values: ["v"] });
    expect(() => Table.remove(data, "Src")).toThrow(ExcelPivotTableError);
    expect(() => Workbook.removeWorksheet(wb, data)).toThrow(ExcelPivotTableError);
    expect(Table.list(data)).toHaveLength(1);
  });

  it("a pivot table on the sheet being removed goes with it", async () => {
    const { wb, data } = book();
    Pivot.add(data, { sourceSheet: data, rows: ["Region"], values: ["Sales"], ref: "E3" });
    expect(Workbook.removeWorksheet(wb, data)).toBe(true);
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    expect(Workbook.listWorksheets(back).map(ws => Worksheet.getName(ws))).toEqual(["Pivot"]);
  });
});

describe("Pivot.add", () => {
  it("refuses a second pivot table over the first one's cells", () => {
    const { data, out } = book();
    Pivot.add(out, { sourceSheet: data, rows: ["Region"], values: ["Sales"] });
    expect(() =>
      Pivot.add(out, { sourceSheet: data, rows: ["Region"], values: ["Sales"] })
    ).toThrow(ExcelPivotTableError);
    expect(() =>
      Pivot.add(out, { sourceSheet: data, rows: ["Region"], values: ["Sales"], ref: "B4" })
    ).toThrow(ExcelPivotTableError);
    Pivot.add(out, { sourceSheet: data, rows: ["Region"], values: ["Sales"], ref: "E3" });
    expect(Pivot.list(out)).toHaveLength(2);
  });

  it.each([
    null,
    { sourceSheet: undefined, rows: 42, values: ["Sales"] },
    { rows: ["Region"] },
    { rows: [1], values: ["Sales"] }
  ])("refuses the model %j with a typed error", model => {
    const { out } = book();
    expect(() => Pivot.add(out, model as never)).toThrow(ExcelArgumentError);
    expect(Pivot.list(out)).toEqual([]);
  });

  it("takes read-only arrays", () => {
    const { data, out } = book();
    const rows = ["Region"] as const;
    const values = ["Sales"] as const;
    expect(Pivot.add(out, { sourceSheet: data, rows, values }).rows).toEqual([0]);
  });
});
