/**
 * Structural edits keep the workbook consistent: after rows or columns are inserted or deleted, a sheet is renamed or
 * removed, everything that names cells by position follows them, as Excel does. Each of these was left stale before.
 */
import {
  Cell,
  Chart,
  Column,
  ConditionalFormatting,
  Pivot,
  DataValidation,
  DefinedNames,
  Row,
  Sparkline,
  Table,
  Workbook,
  Worksheet
} from "@excel";
import { describe, expect, it } from "vitest";

function rangesOf(wb: Workbook.Handle, name: string): string[] {
  return DefinedNames.list(Workbook.getDefinedNames(wb)).find(entry => entry.name === name)!.ranges;
}

function seriesValues(chart: ReturnType<typeof Chart.get>[number], index = 0): string | undefined {
  const series = Chart.getSeries(chart, index) as { val?: { numRef?: { formula?: string } } };
  return series?.val?.numRef?.formula;
}

function seriesCategories(
  chart: ReturnType<typeof Chart.get>[number],
  index = 0
): string | undefined {
  const series = Chart.getSeries(chart, index) as {
    cat?: { strRef?: { formula?: string }; numRef?: { formula?: string } };
  };
  return series?.cat?.strRef?.formula ?? series?.cat?.numRef?.formula;
}

/** A sheet `D` with a table at A3:B5 (header, two rows), and a few things that point into it. */
function build() {
  const wb = Workbook.create();
  const d = Workbook.addWorksheet(wb, "D");
  const other = Workbook.addWorksheet(wb, "Other");
  Table.add(d, {
    name: "Sales",
    ref: "A3",
    columns: [{ name: "Item" }, { name: "Amount" }],
    rows: [
      ["a", 1],
      ["b", 2]
    ]
  });
  Cell.setValue(d, "D1", { formula: "SUM(B4:B5)" });
  Cell.setValue(other, "A1", { formula: "D!B4*2" });
  return { wb, d, other };
}

describe("Row.insert / Row.insertMany", () => {
  it("moves a table with its cells", () => {
    const { d } = build();
    Row.insert(d, 1, ["new"]);
    const table = Table.get(d, "Sales")!;
    expect(Table.getRef(table)).toBe("A4:B6");
    expect(Table.getModel(table).tableRef).toBe("A4:B6");
    expect(Table.getModel(table).rows).toEqual([
      ["a", 1],
      ["b", 2]
    ]);
  });

  it("rewrites formulas on the sheet and on other sheets", () => {
    const { d, other } = build();
    Row.insert(d, 1, ["new"]);
    expect(Cell.getFormula(d, "D2")).toBe("SUM(B5:B6)");
    expect(Cell.getFormula(other, "A1")).toBe("D!B5*2");
  });

  it("grows a table, a formula's range and a name when rows go strictly inside them", () => {
    const { wb, d } = build();
    DefinedNames.add(Workbook.getDefinedNames(wb), "D!$B$4:$B$5", "Amounts");
    Row.insertMany(d, 5, [
      ["c", 3],
      ["d", 4]
    ]);
    const table = Table.get(d, "Sales")!;
    expect(Table.getModel(table).tableRef).toBe("A3:B7");
    expect(Table.getModel(table).rows).toEqual([
      ["a", 1],
      ["c", 3],
      ["d", 4],
      ["b", 2]
    ]);
    expect(Cell.getFormula(d, "D1")).toBe("SUM(B4:B7)");
    expect(rangesOf(wb, "Amounts")).toEqual(["D!$B$4:$B$7"]);
  });

  it("does not grow a table from a row inserted just below it", () => {
    const { d } = build();
    Row.insert(d, 6, ["after"]);
    expect(Table.getModel(Table.get(d, "Sales")!).tableRef).toBe("A3:B5");
  });

  it("puts rows inserted before the totals row into the table", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Table.add(ws, {
      name: "T",
      ref: "A1",
      totalsRow: true,
      columns: [{ name: "x", totalsRowFunction: "sum" }, { name: "y" }],
      rows: [[1, 2]]
    });
    Row.insert(ws, 3, [5, 6]);
    const model = Table.getModel(Table.get(ws, "T")!);
    expect(model.tableRef).toBe("A1:B4");
    expect(model.rows).toEqual([
      [1, 2],
      [5, 6]
    ]);
    expect(Cell.getFormula(ws, "A4")).toBe("SUBTOTAL(109,T[x])");
  });

  it("moves conditional formats, validations, sparklines and chart series", () => {
    const { d } = build();
    ConditionalFormatting.add(d, {
      ref: "B4:B5",
      rules: [{ type: "expression", formulae: ["B4>$D$1"], style: { font: { bold: true } } }]
    });
    DataValidation.add(d.dataValidations, "B4:B5", {
      type: "whole",
      operator: "greaterThan",
      formulae: ["$D$1"]
    });
    Cell.setValidation(d, "A4", { type: "list", formulae: ["$A$4:$A$5"] });
    Sparkline.add(d, { type: "line", sparklines: [{ dataRef: "D!B4:B5", cellRef: "C4" }] });
    Chart.addColumn(
      d,
      { series: [{ name: "Amount", categories: "D!$A$4:$A$5", values: "D!$B$4:$B$5" }] },
      { tl: { col: 5, row: 1 }, br: { col: 10, row: 10 } }
    );

    Row.insert(d, 2, []);
    const chart = Chart.get(d)[0];

    expect(ConditionalFormatting.list(d)).toMatchObject([
      { ref: "B5:B6", rules: [{ formulae: ["B5>$D$1"] }] }
    ]);
    expect(Cell.getValidation(d, "B6")).toMatchObject({ formulae: ["$D$1"] });
    expect(Cell.getValidation(d, "B4")).toBeUndefined();
    expect(Cell.getValidation(d, "A5")).toMatchObject({ formulae: ["$A$5:$A$6"] });
    expect(Cell.getValidation(d, "A4")).toBeUndefined();
    expect(Sparkline.list(d)[0].sparklines).toEqual([{ dataRef: "D!B5:B6", cellRef: "C5" }]);
    expect(seriesValues(chart)).toBe("D!$B$5:$B$6");
    expect(seriesCategories(chart)).toBe("D!$A$5:$A$6");
  });

  it("keeps a filled (shared) formula shared, pointing at its master's new place", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.addMany(ws, [[1], [2], [3], [4]]);
    Worksheet.fillFormula(ws, "B1:B4", "A1*2");
    Row.insert(ws, 1, []);
    expect(Cell.getFormula(ws, "B2")).toBe("A2*2");
    expect(Cell.getFormula(ws, "B5")).toBe("A5*2");
    expect(Cell.getValue(ws, "B5")).toMatchObject({ sharedFormula: "B2" });
  });

  it("unshares a filled formula whose master is deleted, keeping what each member computes", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.addMany(ws, [[1], [2], [3], [4]]);
    Worksheet.fillFormula(ws, "B1:B4", "A1*2+$A$4");
    Row.splice(ws, 1, 1);
    expect(Cell.getFormula(ws, "B1")).toBe("A1*2+$A$3");
    expect(Cell.getFormula(ws, "B3")).toBe("A3*2+$A$3");
    expect(Cell.getValue(ws, "B3")).not.toHaveProperty("sharedFormula");
  });

  it("keeps a member pointing at a plain formula it is shared from", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "B1", { formula: "A1*2" });
    Cell.setValue(ws, "B2", { sharedFormula: "B1" });
    Row.insert(ws, 1, []);
    expect(Cell.getFormula(ws, "B2")).toBe("A2*2");
    expect(Cell.getFormula(ws, "B3")).toBe("A3*2");
  });

  it("moves the auto filter, print area, print titles, comments and ignored errors", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    ws.autoFilter = "A2:C5";
    ws.pageSetup.printArea = "A2:C5&&E2:F3";
    ws.pageSetup.printTitlesRow = "2:2";
    ws.threadedComments.push({ ref: "B3", comment: { text: "hi" } as never });
    ws.ignoredErrors.push({ ref: "A3:A5 C9", numberStoredAsText: true });
    Row.insert(ws, 1, []);
    expect(ws.autoFilter).toBe("A3:C6");
    expect(ws.pageSetup.printArea).toBe("A3:C6&&E3:F4");
    expect(ws.pageSetup.printTitlesRow).toBe("3:3");
    expect(ws.threadedComments[0].ref).toBe("B4");
    expect(ws.ignoredErrors[0].ref).toBe("A4:A6 C10");
  });

  it("moves a pivot table's source range", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Data");
    Row.addMany(ws, [
      ["k", "v"],
      ["a", 1],
      ["b", 2]
    ]);
    const report = Workbook.addWorksheet(wb, "Report");
    Pivot.add(report, { sourceSheet: ws, rows: ["k"], values: ["v"] });
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    const data = Workbook.getWorksheet(back, "Data")!;
    Row.insert(data, 1, ["title"]);
    Worksheet.setName(data, "Data 2");
    const pivot = Pivot.list(Workbook.getWorksheet(back, "Report")!)[0];
    expect(pivot.cacheDefinition).toMatchObject({ sourceRef: "A2:B4", sourceSheet: "Data 2" });
  });

  it("moves an array formula and its range", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "C2", { formula: "A2:A3*2", shareType: "array", ref: "C2:C3" });
    Row.insert(ws, 1, []);
    expect(Cell.getValue(ws, "C3")).toMatchObject({ formula: "A3:A4*2", ref: "C3:C4" });
  });
});

describe("deleting rows", () => {
  it("turns references into the deleted rows into #REF! and shrinks ranges across them", () => {
    const { wb, d, other } = build();
    DefinedNames.add(Workbook.getDefinedNames(wb), "D!$B$4", "First");
    Row.splice(d, 4, 1);
    expect(Cell.getFormula(d, "D1")).toBe("SUM(B4:B4)");
    expect(Cell.getFormula(other, "A1")).toBe("#REF!*2");
    expect(DefinedNames.list(Workbook.getDefinedNames(wb))).toContainEqual(
      expect.objectContaining({ name: "First", formulaExpression: "#REF!" })
    );
    const table = Table.get(d, "Sales")!;
    expect(Table.getModel(table).tableRef).toBe("A3:B4");
    expect(Table.getModel(table).rows).toEqual([["b", 2]]);
  });

  it("removes a table whose rows are all deleted", () => {
    const { wb, d } = build();
    Row.splice(d, 2, 5);
    expect(Table.get(d, "Sales")).toBeUndefined();
    // the name is free again
    Table.add(d, { name: "Sales", ref: "H1", columns: [{ name: "x" }], rows: [[1]] });
    expect(Workbook.getWorksheet(wb, "D")).toBe(d);
  });

  it("drops conditional formats and validations whose cells are all deleted", () => {
    const { d } = build();
    ConditionalFormatting.add(d, {
      ref: "B4",
      rules: [{ type: "expression", formulae: ["TRUE"], style: { font: { bold: true } } }]
    });
    Cell.setValidation(d, "B4", { type: "whole", operator: "greaterThan", formulae: [0] });
    Row.splice(d, 4, 1);
    expect(ConditionalFormatting.list(d)).toEqual([]);
    expect(Cell.getValidation(d, "B4")).toBeUndefined();
  });

  it("keeps references to rows a splice overwrites", () => {
    const { d } = build();
    Cell.setValue(d, "F1", { formula: "A4&A5&A6" });
    Row.splice(d, 4, 2, ["x", 9]);
    expect(Cell.getFormula(d, "F1")).toBe("A4&#REF!&A5");
  });
});

describe("Column.splice", () => {
  it("moves a table, formulas and names right when a column is inserted before them", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Table.add(ws, {
      name: "T",
      ref: "B1",
      columns: [{ name: "x" }, { name: "y" }],
      rows: [[1, 2]]
    });
    Cell.setValue(ws, "E1", { formula: "SUM(B2:C2)+SUM(B:B)" });
    DefinedNames.add(Workbook.getDefinedNames(wb), "S!$B$2:$C$2", "Second");
    Column.splice(ws, 1, 0, []);
    expect(Table.getRef(Table.get(ws, "T")!)).toBe("C1:D2");
    expect(Cell.getFormula(ws, "F1")).toBe("SUM(C2:D2)+SUM(C:C)");
    expect(rangesOf(wb, "Second")).toEqual(["S!$C$2:$D$2"]);
  });

  it("gives a column inserted inside a table a definition and a header", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Table.add(ws, {
      name: "T",
      ref: "A1",
      columns: [{ name: "x" }, { name: "Column1" }],
      rows: [[1, 2]]
    });
    Column.splice(ws, 2, 0, []);
    const model = Table.getModel(Table.get(ws, "T")!);
    expect(model.tableRef).toBe("A1:C2");
    expect(model.columns.map(c => c.name)).toEqual(["x", "Column2", "Column1"]);
    expect(Cell.getValue(ws, "B1")).toBe("Column2");
    expect(model.rows).toEqual([[1, null, 2]]);
  });

  it("shrinks a table when one of its columns is deleted", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Table.add(ws, {
      name: "T",
      ref: "A1",
      columns: [{ name: "x" }, { name: "y" }, { name: "z" }],
      rows: [[1, 2, 3]]
    });
    Column.splice(ws, 2, 1);
    const model = Table.getModel(Table.get(ws, "T")!);
    expect(model.tableRef).toBe("A1:B2");
    expect(model.columns.map(c => c.name)).toEqual(["x", "z"]);
    expect(model.rows).toEqual([[1, 3]]);
  });
});

describe("Worksheet.setName", () => {
  it("renames the sheet in formulas, names and chart series everywhere", () => {
    const { wb, d, other } = build();
    DefinedNames.add(Workbook.getDefinedNames(wb), "D!$B$4:$B$5", "Amounts");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Twice", "D!$B$4*2");
    Chart.addColumn(
      other,
      { series: [{ name: "Amount", values: "D!$B$4:$B$5" }] },
      { tl: { col: 5, row: 1 }, br: { col: 10, row: 10 } }
    );
    Cell.setValue(other, "A2", { formula: '"D!B4"&d!B5' });

    Worksheet.setName(d, "Q1 Data");
    const chart = Chart.get(other)[0];

    expect(Cell.getFormula(other, "A1")).toBe("'Q1 Data'!B4*2");
    expect(Cell.getFormula(other, "A2")).toBe(`"D!B4"&'Q1 Data'!B5`);
    expect(Cell.getFormula(d, "D1")).toBe("SUM(B4:B5)");
    expect(rangesOf(wb, "Amounts")).toEqual(["'Q1 Data'!$B$4:$B$5"]);
    expect(rangesOf(wb, "Twice")).toEqual(["'Q1 Data'!$B$4*2"]);
    expect(seriesValues(chart)).toBe("'Q1 Data'!$B$4:$B$5");
  });

  it("changes nothing when the new name is refused", () => {
    const { d, other } = build();
    expect(() => Worksheet.setName(d, "Other")).toThrow();
    expect(Cell.getFormula(other, "A1")).toBe("D!B4*2");
  });
});

describe("Workbook.removeWorksheet", () => {
  it("turns references to the removed sheet into #REF!", () => {
    const { wb, other } = build();
    DefinedNames.add(Workbook.getDefinedNames(wb), "D!$B$4", "First");
    Cell.setValue(other, "A2", { formula: "Other!A1+1" });
    Workbook.removeWorksheet(wb, "D");
    expect(Cell.getFormula(other, "A1")).toBe("#REF!*2");
    expect(Cell.getFormula(other, "A2")).toBe("Other!A1+1");
    expect(DefinedNames.list(Workbook.getDefinedNames(wb))).toContainEqual(
      expect.objectContaining({ name: "First", formulaExpression: "#REF!" })
    );
  });
});

describe("round trip", () => {
  it.each([
    ["xlsx", false],
    ["xlsb", false],
    ["xlsx", true]
  ] as const)(
    "a %s written after the edits reads back pointing at the same data (edited after a read: %s)",
    async (format, reread) => {
      let { wb, d, other } = build();
      DefinedNames.add(Workbook.getDefinedNames(wb), "D!$B$4:$B$5", "Amounts");
      ConditionalFormatting.add(d, {
        ref: "B4:B5",
        rules: [{ type: "expression", formulae: ["B4>1"], style: { font: { bold: true } } }]
      });
      Cell.setValidation(d, "B4", { type: "whole", operator: "greaterThan", formulae: [0] });
      Chart.addColumn(
        other,
        { series: [{ name: "Amount", values: "D!$B$4:$B$5" }] },
        { tl: { col: 5, row: 1 }, br: { col: 10, row: 10 } }
      );

      if (reread) {
        // Edited after a read, a chart is patched in the bytes it was read from rather than rebuilt.
        wb = await Workbook.read(await Workbook.toBuffer(wb));
        d = Workbook.getWorksheet(wb, "D")!;
        other = Workbook.getWorksheet(wb, "Other")!;
      }

      Row.insert(d, 1, ["top"]);
      Column.splice(d, 1, 0, ["left"]);
      Worksheet.setName(d, "Data 2");

      const back = await Workbook.read(await Workbook.toBuffer(wb, { format }));
      const d2 = Workbook.getWorksheet(back, "Data 2")!;
      const o2 = Workbook.getWorksheet(back, "Other")!;
      expect(Table.getModel(Table.get(d2, "Sales")!).tableRef).toBe("B4:C6");
      expect(Cell.getValue(d2, "C5")).toBe(1);
      expect(Cell.getFormula(d2, "E2")).toBe("SUM(C5:C6)");
      expect(Cell.getFormula(o2, "A1")).toBe("'Data 2'!C5*2");
      expect(rangesOf(back, "Amounts")).toEqual(["'Data 2'!$C$5:$C$6"]);
      expect(ConditionalFormatting.list(d2)).toMatchObject([
        { ref: "C5:C6", rules: [{ formulae: ["C5>1"] }] }
      ]);
      expect(Cell.getValidation(d2, "C5")).toMatchObject({ type: "whole" });
      if (format === "xlsx") {
        // The XLSB reader carries a chart's part through without modelling it.
        expect(seriesValues(Chart.get(o2)[0])).toBe("'Data 2'!$C$5:$C$6");
      }
    }
  );
});
