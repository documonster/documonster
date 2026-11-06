/**
 * A table's edits reach the sheet as they are made, and the sheet's cells are what a table holds: a value edited
 * through `Cell.setValue` survives the table's next edit instead of being overwritten by the copy the table took when
 * it was built.
 */
import { Cell, ExcelArgumentError, Table, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function setup() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const t = Table.add(ws, {
    name: "T",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "A" }, { name: "B" }],
    rows: [
      [1, 2],
      [3, 4]
    ]
  });
  return { wb, ws, t };
}

async function reread(wb: Workbook.Handle) {
  const back = await Workbook.read(await Workbook.toBuffer(wb));
  return Workbook.getWorksheet(back, "S")!;
}

describe("a cell edited through Cell.setValue survives the table's next edit", () => {
  it("addRow", async () => {
    const { wb, ws, t } = setup();
    Cell.setValue(ws, "A2", 99);
    Table.addRow(t, [7, 8]);
    expect(Cell.getValue(ws, "A2")).toBe(99);
    expect(Cell.getValue(ws, "A4")).toBe(7);
    expect(Cell.getValue(await reread(wb), "A2")).toBe(99);
  });

  it("removeRows, addColumn, removeColumns", () => {
    const { ws, t } = setup();
    Cell.setValue(ws, "B3", "edited");
    Table.removeRows(t, 0);
    expect(Cell.getValue(ws, "B2")).toBe("edited");
    Table.addColumn(t, { name: "C" }, [5]);
    expect(Cell.getValue(ws, "B2")).toBe("edited");
    expect(Cell.getValue(ws, "C2")).toBe(5);
    Cell.setValue(ws, "C2", 6);
    Table.removeColumns(t, 0);
    expect([Cell.getValue(ws, "A2"), Cell.getValue(ws, "B2")]).toEqual(["edited", 6]);
  });

  it("setRef moves the edited value with the table", () => {
    const { ws, t } = setup();
    Cell.setValue(ws, "B2", "moved");
    Table.setRef(t, "D5");
    expect(Cell.getValue(ws, "E6")).toBe("moved");
    expect(Cell.getValue(ws, "B2")).toBeNull();
  });

  it("setColumnStyle keeps the values", () => {
    const { ws, t } = setup();
    Cell.setValue(ws, "A3", "kept");
    Table.setColumnStyle(Table.getColumn(t, 0)!, { font: { bold: true } });
    expect(Cell.getValue(ws, "A3")).toBe("kept");
    expect(Cell.getFont(ws, "A3")?.bold).toBe(true);
  });
});

describe("edits apply as they are made", () => {
  it("addColumn and setRef without commit leave the sheet and the file consistent", async () => {
    const { wb, ws, t } = setup();
    Table.addColumn(t, { name: "C" }, [5, 6]);
    expect([Cell.getValue(ws, "C1"), Cell.getValue(ws, "C2")]).toEqual(["C", 5]);
    Table.setRef(t, "E5");
    expect(Cell.getValue(ws, "A1")).toBeNull();
    expect([Cell.getValue(ws, "E5"), Cell.getValue(ws, "G7")]).toEqual(["A", 6]);
    const back = await reread(wb);
    expect([Cell.getValue(back, "A1"), Cell.getValue(back, "E5")]).toEqual([null, "A"]);
  });

  it("column renames and totals settings reach the header and totals rows", () => {
    const { ws, t } = setup();
    Table.setTotalsRow(t, true);
    Table.setColumnName(Table.getColumn(t, 1)!, "Amount");
    Table.setColumnTotalsRowFunction(Table.getColumn(t, 1)!, "sum");
    expect(Cell.getValue(ws, "B1")).toBe("Amount");
    expect(Cell.getFormula(ws, "B4")).toBe("SUBTOTAL(109,T[Amount])");
    Table.setName(t, "Sales");
    expect(Cell.getFormula(ws, "B4")).toBe("SUBTOTAL(109,Sales[Amount])");
    expect(Cell.getValue(ws, "B3")).toBe(4);
  });

  it("{ commit: false } no longer defers: the row is on the sheet, and in the written file", async () => {
    const { wb, ws, t } = setup();
    Table.addRow(t, [5, 6], undefined, { commit: false });
    expect(Cell.getValue(ws, "A4")).toBe(5);
    Cell.setValue(ws, "A2", 999);
    Table.commit(t);
    expect([Cell.getValue(ws, "A2"), Cell.getValue(ws, "A4")]).toEqual([999, 5]);
    const back = await reread(wb);
    expect(Table.getModel(Table.list(back)[0]).tableRef).toBe("A1:B4");
    expect(Cell.getValue(back, "A4")).toBe(5);
  });
});

describe("a cell's own formatting stays with its value", () => {
  function styled() {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const t = Table.add(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      columns: [{ name: "A", style: { font: { italic: true } } }, { name: "B" }],
      rows: [
        [1, 2],
        [3, 4]
      ]
    });
    // Overrides the column's own facet, and adds one the column does not set.
    Cell.setFont(ws, "A2", { bold: true });
    Cell.setFill(ws, "B3", { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } });
    return { ws, t };
  }

  it("appending a row leaves existing cells' formatting alone and styles the new row from the column", () => {
    const { ws, t } = styled();
    Table.addRow(t, [5, 6]);
    expect(Cell.getFont(ws, "A2")).toEqual({ bold: true });
    expect((Cell.getFill(ws, "B3") as { fgColor?: unknown } | undefined)?.fgColor).toEqual({
      argb: "FFFFFF00"
    });
    expect(Cell.getFont(ws, "A4")).toEqual({ italic: true });
  });

  it("inserting a row moves the formatting down with the values", () => {
    const { ws, t } = styled();
    Table.addRow(t, [0, 0], 0);
    expect(Cell.getValue(ws, "A3")).toBe(1);
    expect(Cell.getFont(ws, "A3")).toEqual({ bold: true });
    expect(Cell.getFont(ws, "A2")).toEqual({ italic: true });
    expect((Cell.getFill(ws, "B4") as { fgColor?: unknown } | undefined)?.fgColor).toEqual({
      argb: "FFFFFF00"
    });
    expect(Cell.getFill(ws, "B3")).toBeUndefined();
  });

  it("inserting and removing columns, and moving the table, carry it too", () => {
    const { ws, t } = styled();
    Table.addColumn(t, { name: "First" }, [10, 30], 0);
    expect(Cell.getFont(ws, "B2")).toEqual({ bold: true });
    expect((Cell.getFill(ws, "C3") as { fgColor?: unknown } | undefined)?.fgColor).toEqual({
      argb: "FFFFFF00"
    });
    Table.removeColumns(t, 0);
    expect(Cell.getFont(ws, "A2")).toEqual({ bold: true });
    Table.setRef(t, "D4");
    expect(Cell.getValue(ws, "D5")).toBe(1);
    expect(Cell.getFont(ws, "D5")).toEqual({ bold: true });
    expect((Cell.getFill(ws, "E6") as { fgColor?: unknown } | undefined)?.fgColor).toEqual({
      argb: "FFFFFF00"
    });
  });

  it("setting the column's style is what restyles its cells", () => {
    const { ws, t } = styled();
    Table.setColumnStyle(Table.getColumn(t, 0)!, { font: { underline: true } });
    expect(Cell.getFont(ws, "A2")).toEqual({ underline: true });
    expect((Cell.getFill(ws, "B3") as { fgColor?: unknown } | undefined)?.fgColor).toEqual({
      argb: "FFFFFF00"
    });
  });

  it("a moved cell does not share its style with the cell it left", () => {
    const { ws, t } = styled();
    Table.setRef(t, "D4");
    Cell.setFont(ws, "D5", { strike: true });
    expect(Cell.getFont(ws, "A2")).toEqual({ bold: true });
  });
});

describe("the model reports what the sheet holds", () => {
  it("getModel rows, and the rows a pivot or chart reads, follow Cell.setValue", () => {
    const { ws, t } = setup();
    Cell.setValue(ws, "B3", 40);
    expect(Table.getModel(t).rows).toEqual([
      [1, 2],
      [3, 40]
    ]);
  });

  it("reading the model does not rewrite the rows the caller passed in", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const rows = [[1], [2]];
    const t = Table.add(ws, { name: "T", ref: "A1", columns: [{ name: "A" }], rows });
    Cell.setValue(ws, "A2", 9);
    expect(Table.getModel(t).rows).toEqual([[9], [2]]);
    expect(rows).toEqual([[1], [2]]);
  });
});

describe("getModel and setModel agree", () => {
  it("a model read, edited and set back is what the table holds", () => {
    const { ws, t } = setup();
    const model = Table.getModel(t);
    model.rows[0][0] = 777;
    Table.setModel(t, model);
    expect(Cell.getValue(ws, "A2")).toBe(777);
    expect(Table.getModel(t).rows[0][0]).toBe(777);
  });

  it("getModel is a copy: editing it changes nothing until it is set", () => {
    const { ws, t } = setup();
    const model = Table.getModel(t);
    model.name = "Other";
    model.columns[0].name = "Z";
    model.rows[0][0] = 5;
    expect(Table.getName(t)).toBe("T");
    expect(Cell.getValue(ws, "A1")).toBe("A");
    expect(Cell.getValue(ws, "A2")).toBe(1);
  });

  it("setModel renames through the workbook, so the name stays unique", () => {
    const { ws, t } = setup();
    Table.setModel(t, { ...Table.getModel(t), name: "Renamed" });
    expect(Table.get(ws, "Renamed")).toBe(t);
    expect(Table.get(ws, "T")).toBeUndefined();
    expect(() =>
      Table.add(ws, { name: "Renamed", ref: "E1", columns: [{ name: "A" }], rows: [[1]] })
    ).toThrow(/already exists/);
    const other = Table.add(ws, {
      name: "Other",
      ref: "E1",
      columns: [{ name: "A" }],
      rows: [[1]]
    });
    expect(() => Table.setModel(other, { ...Table.getModel(other), name: "renamed" })).toThrow(
      /already exists/
    );
    expect(Table.getName(other)).toBe("Other");
  });

  it("setModel moves and reshapes the table, clearing what it no longer covers", () => {
    const { ws, t } = setup();
    Table.setModel(t, { ...Table.getModel(t), ref: "D1", rows: [[9, 8]] });
    expect([Cell.getValue(ws, "A1"), Cell.getValue(ws, "A2"), Cell.getValue(ws, "A3")]).toEqual([
      null,
      null,
      null
    ]);
    expect([Cell.getValue(ws, "D1"), Cell.getValue(ws, "D2"), Cell.getValue(ws, "E2")]).toEqual([
      "A",
      9,
      8
    ]);
    expect(Table.getModel(t).tableRef).toBe("D1:E2");
  });

  it("the deprecated Table.model returns the model the table uses, rows read from the sheet", () => {
    const { ws, t } = setup();
    Cell.setValue(ws, "B3", 40);
    const live = Table.model(t);
    expect(live).toBe(Table.model(t));
    expect(live.rows).toEqual([
      [1, 2],
      [3, 40]
    ]);
  });
});

describe("an edit that throws leaves the table as it was", () => {
  it("setRef to an invalid address", () => {
    const { ws, t } = setup();
    expect(() => Table.setRef(t, "not an address")).toThrow();
    expect(Table.getRef(t)).toBe("A1:B3");
    Table.addRow(t, [5, 6]);
    expect(Cell.getValue(ws, "A4")).toBe(5);
  });

  it("an empty column name, and a column without one", () => {
    const { ws, t } = setup();
    expect(() => Table.setColumnName(Table.getColumn(t, 0)!, "")).toThrow(/must have a name/);
    expect(Table.getColumnName(Table.getColumn(t, 0)!)).toBe("A");
    expect(() => Table.addColumn(t, { name: "" }, [1, 2])).toThrow(/must have a name/);
    expect(Table.getModel(t).columns.map(c => c.name)).toEqual(["A", "B"]);
    expect(Cell.getValue(ws, "C1")).toBeNull();
  });

  it("getColumn answers undefined for an index the table has no column at; the deprecated column hands out an empty view", () => {
    const { t } = setup();
    expect(Table.getColumn(t, 2)).toBeUndefined();
    expect(Table.getColumn(t, -1)).toBeUndefined();
    expect(() => Table.getColumn(t, 0.5)).toThrow(ExcelArgumentError);
    expect(Table.column(t, 2).column).toBeUndefined();
  });
});

describe("clearing and replacing a column's style", () => {
  it("takes back what the column gave its cells and keeps what a caller set", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const t = Table.add(ws, {
      name: "T",
      ref: "A1",
      columns: [{ name: "A", style: { font: { italic: true } } }],
      rows: [[1], [2]]
    });
    Cell.setFont(ws, "A3", { bold: true });
    Table.setColumnStyle(Table.getColumn(t, 0)!, undefined);
    expect(Cell.getFont(ws, "A2")).toBeUndefined();
    expect(Cell.getFont(ws, "A3")).toEqual({ bold: true });
    Table.setColumnStyle(Table.getColumn(t, 0)!, { font: { underline: true } });
    expect(Cell.getFont(ws, "A2")).toEqual({ underline: true });
  });

  it("a column inserted before others takes its own style, not the one it displaced", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const t = Table.add(ws, {
      name: "T",
      ref: "A1",
      columns: [{ name: "A", style: { font: { italic: true } } }],
      rows: [[1]]
    });
    Table.addColumn(t, { name: "New" }, [7], 0);
    expect(Cell.getFont(ws, "A2")).toBeUndefined();
    expect(Cell.getFont(ws, "B2")).toEqual({ italic: true });
  });
});

describe("appending rows costs only the new cells", () => {
  it("a thousand appends to a 20-column table stay linear", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const columns = Array.from({ length: 20 }, (_, i) => ({ name: `C${i}` }));
    const t = Table.add(ws, { name: "T", ref: "A1", columns, rows: [columns.map(() => 0)] });
    const started = performance.now();
    for (let i = 0; i < 1000; i++) {
      Table.addRow(
        t,
        columns.map(() => i)
      );
    }
    // Rewriting the table per append made this quadratic: ~20M cell writes for the loop.
    expect(performance.now() - started).toBeLessThan(5000);
    expect(Cell.getValue(ws, "A1002")).toBe(999);
    expect(Table.getModel(t).tableRef).toBe("A1:T1002");
  });
});
