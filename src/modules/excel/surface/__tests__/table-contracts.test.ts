/**
 * Table edits check what they are given before anything on the sheet changes: positions, values, the area against the
 * sheet's other tables, names against the workbook, and the handle against removal.
 */
import { Cell, ExcelArgumentError, Row, Table, TableError, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

function setup() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const t = Table.add(ws, {
    name: "Sales",
    ref: "A1",
    columns: [{ name: "a" }, { name: "b" }],
    rows: [
      [1, 2],
      [3, 4]
    ]
  });
  return { wb, ws, t };
}

function cells(ws: ReturnType<typeof setup>["ws"], addresses: string[]): unknown[] {
  return addresses.map(address => Cell.getValue(ws, address));
}

describe("row and column positions", () => {
  it.each([99, -1, 1.5, Number.NaN, 3])("addRow refuses position %s and changes nothing", at => {
    const { ws, t } = setup();
    expect(() => Table.addRow(t, [5, 6], at)).toThrow(TableError);
    expect(Table.getModel(t).rows).toEqual([
      [1, 2],
      [3, 4]
    ]);
    expect(cells(ws, ["A2", "A3", "A4"])).toEqual([1, 3, null]);
  });

  it("addRow accepts every position from 0 to the row count", () => {
    const { t } = setup();
    Table.addRow(t, [0, 0], 0);
    Table.addRow(t, [9, 9], 3);
    expect(Table.getModel(t).rows).toEqual([
      [0, 0],
      [1, 2],
      [3, 4],
      [9, 9]
    ]);
  });

  it("addRow refuses more values than columns, and values no cell can hold, before writing", () => {
    const { ws, t } = setup();
    expect(() => Table.addRow(t, [5, 6, 7])).toThrow(/2 column/);
    expect(() => Table.addRow(t, [5, Number.NaN])).toThrow();
    expect(() => Table.addRow(t, null as never)).toThrow(TableError);
    expect(cells(ws, ["A4", "B4", "C4"])).toEqual([null, null, null]);
    expect(Table.getRef(t)).toBe("A1:B3");
  });

  it.each([
    [-1, 1],
    [1.5, 1],
    [2, 1],
    [0, -1],
    [0, Number.NaN]
  ])("removeRows refuses (%s, %s)", (index, count) => {
    const { t } = setup();
    expect(() => Table.removeRows(t, index, count)).toThrow(TableError);
    expect(Table.getModel(t).rows).toHaveLength(2);
  });

  it("removeRows removes up to the end", () => {
    const { t } = setup();
    Table.removeRows(t, 1, 10);
    expect(Table.getModel(t).rows).toEqual([[1, 2]]);
  });

  it.each([
    [-1, 1],
    [2, 1],
    [0.5, 1],
    [0, -1]
  ])("removeColumns refuses (%s, %s)", (index, count) => {
    const { t } = setup();
    expect(() => Table.removeColumns(t, index, count)).toThrow(TableError);
    expect(Table.getModel(t).columns.map(column => column.name)).toEqual(["a", "b"]);
  });

  it("removeColumns keeps at least one column", () => {
    const { t } = setup();
    expect(() => Table.removeColumns(t, 0, 2)).toThrow(/at least one column/);
    expect(Table.getModel(t).columns).toHaveLength(2);
  });

  it.each([99, -1, Number.NaN])("addColumn refuses position %s", at => {
    const { ws, t } = setup();
    expect(() => Table.addColumn(t, { name: "c" }, [7, 8], at)).toThrow(TableError);
    expect(Cell.getValue(ws, "C1")).toBeNull();
  });

  it("addColumn refuses more values than data rows", () => {
    const { ws, t } = setup();
    expect(() => Table.addColumn(t, { name: "c" }, [7, 8, 9])).toThrow(/2 data row/);
    expect(Cell.getValue(ws, "C1")).toBeNull();
  });
});

describe("tables cannot overlap", () => {
  it("Table.add refuses an area another table covers, and writes nothing", () => {
    const { ws } = setup();
    Cell.setValue(ws, "A5", "keep");
    expect(() =>
      Table.add(ws, { name: "T_b", ref: "A2", columns: [{ name: "x" }], rows: [[9]] })
    ).toThrow(/overlaps table "Sales"/);
    expect(cells(ws, ["A2", "A3"])).toEqual([1, 3]);
    expect(Table.list(ws)).toHaveLength(1);
  });

  it("addRow refuses to grow into another table", () => {
    const { ws, t } = setup();
    Table.add(ws, { name: "T_c", ref: "A5", columns: [{ name: "c" }], rows: [[5]] });
    Table.addRow(t, [5, 6]);
    expect(() => Table.addRow(t, [7, 8])).toThrow(/overlaps table "T_c"/);
    expect(cells(ws, ["A4", "A5", "A6"])).toEqual([5, "c", 5]);
  });

  it("setRef refuses to move onto another table", () => {
    const { ws, t } = setup();
    Table.add(ws, { name: "T_d", ref: "D1", columns: [{ name: "d" }], rows: [[5]] });
    expect(() => Table.setRef(t, "C1")).toThrow(TableError);
    expect(Table.getRef(t)).toBe("A1:B3");
    expect(Cell.getValue(ws, "D1")).toBe("d");
  });
});

describe("a refused add or setModel leaves the sheet as it was", () => {
  it("a duplicate name", () => {
    const { ws } = setup();
    Cell.setValue(ws, "D1", "keep");
    Cell.setValue(ws, "D2", "keep2");
    expect(() =>
      Table.add(ws, { name: "Sales", ref: "D1", columns: [{ name: "z" }], rows: [[99]] })
    ).toThrow(/already exists/);
    expect(cells(ws, ["D1", "D2"])).toEqual(["keep", "keep2"]);
  });

  it("a value no cell can hold", () => {
    const { ws } = setup();
    Cell.setValue(ws, "D1", "keep");
    Cell.setValue(ws, "D2", "keep2");
    expect(() =>
      Table.add(ws, { name: "T_x", ref: "D1", columns: [{ name: "a" }], rows: [[1], [Number.NaN]] })
    ).toThrow();
    expect(cells(ws, ["D1", "D2", "D3"])).toEqual(["keep", "keep2", null]);
  });

  it("setModel with a value no cell can hold", () => {
    const { ws, t } = setup();
    const before = Table.getModel(t);
    expect(() =>
      Table.setModel(t, {
        ...before,
        rows: [
          [1, 2],
          [Number.NaN, 0]
        ]
      })
    ).toThrow();
    expect(Table.getModel(t)).toEqual(before);
    expect(cells(ws, ["A2", "A3", "A4"])).toEqual([1, 3, null]);
  });

  it("the caller's model is copied, not kept", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const model = { name: "T", ref: "A1", columns: [{ name: "a" }], rows: [[1]] };
    const t = Table.add(ws, model);
    model.columns[0].name = "changed";
    expect(Table.getColumnName(Table.getColumn(t, 0))).toBe("a");
    expect(model).toEqual({ name: "T", ref: "A1", columns: [{ name: "changed" }], rows: [[1]] });
  });
});

describe("a removed table's handle refuses edits", () => {
  it("after Table.remove", () => {
    const { ws, t } = setup();
    const column = Table.getColumn(t, 0);
    expect(Table.remove(ws, t)).toBe(true);
    expect(() => Table.setRef(t, "D10")).toThrow(ExcelArgumentError);
    expect(() => Table.addRow(t, [1, 2])).toThrow(ExcelArgumentError);
    expect(() => Table.setName(t, "Again")).toThrow(ExcelArgumentError);
    expect(() => Table.setDisplayName(t, "Again")).toThrow(ExcelArgumentError);
    expect(() => Table.setTheme(t, "TableStyleLight1")).toThrow(ExcelArgumentError);
    expect(() => Table.setColumnFilterButton(column, false)).toThrow(ExcelArgumentError);
    expect(() => Table.setModel(t, Table.getModel(t))).toThrow(ExcelArgumentError);
    // The cells stay, as when Excel converts a table to a range, and the handle still reads.
    expect(cells(ws, ["A1", "A2"])).toEqual(["a", 1]);
    expect(Table.getName(t)).toBe("Sales");
  });

  it("after a row edit deletes the table", () => {
    const { ws, t } = setup();
    Row.splice(ws, 1, 3);
    expect(Table.list(ws)).toEqual([]);
    expect(() => Table.addRow(t, [1, 2])).toThrow(ExcelArgumentError);
  });

  it("the view of a removed column refuses edits", () => {
    const { t } = setup();
    const column = Table.getColumn(t, 1);
    Table.removeColumns(t, 1);
    expect(() => Table.setColumnName(column, "x")).toThrow(ExcelArgumentError);
    expect(() => Table.setColumnFilterButton(column, false)).toThrow(ExcelArgumentError);
  });

  it("on a removed sheet, setDisplayName is refused too", () => {
    const { wb, ws, t } = setup();
    Workbook.addWorksheet(wb, "Other");
    Workbook.removeWorksheet(wb, ws);
    expect(() => Table.setDisplayName(t, "Other_name")).toThrow(ExcelArgumentError);
  });
});

describe("names and display names are unique in the workbook", () => {
  it("setDisplayName refuses another table's name or display name", () => {
    const { ws, t } = setup();
    const other = Table.add(ws, {
      name: "Other",
      ref: "F1",
      columns: [{ name: "x" }],
      rows: [[1]]
    });
    expect(() => Table.setDisplayName(other, "Sales")).toThrow(/already exists/);
    expect(() => Table.setDisplayName(other, "SALES")).toThrow(/already exists/);
    Table.setDisplayName(t, "Shown");
    expect(() => Table.setDisplayName(other, "shown")).toThrow(/already exists/);
    expect(() => Table.setName(other, "Shown")).toThrow(/already exists/);
    expect(() =>
      Table.add(ws, {
        name: "Third",
        displayName: "Shown",
        ref: "H1",
        columns: [{ name: "y" }],
        rows: [[1]]
      })
    ).toThrow(/already exists/);
    expect(Table.getDisplayName(other)).toBe("Other");
  });
});

describe("Table.getRef is the table's range however the table was made", () => {
  it("built in memory, edited, and read back", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const t = Table.add(ws, {
      name: "T_a",
      ref: "B2",
      columns: [{ name: "a" }, { name: "b" }],
      rows: [[1, 2]]
    });
    expect(Table.getRef(t)).toBe("B2:C3");
    Table.addRow(t, [3, 4]);
    expect(Table.getRef(t)).toBe("B2:C4");
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    expect(Table.getRef(Table.list(Workbook.getWorksheet(back, "S")!)[0])).toBe("B2:C4");
    // The deprecated form keeps answering with the reference as given.
    expect(Table.ref(t)).toBe("B2");
  });
});

describe("Table.get and Table.remove", () => {
  it("remove takes the handle as well as the name", () => {
    const { ws, t } = setup();
    const other = Workbook.addWorksheet(Workbook.create(), "X");
    expect(Table.remove(other, t)).toBe(false);
    expect(Table.remove(ws, t)).toBe(true);
    expect(Table.remove(ws, t)).toBe(false);
  });

  it("refuse an argument that is neither", () => {
    const { ws } = setup();
    expect(() => Table.get(ws, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Table.remove(ws, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Table.remove(ws, {} as never)).toThrow(ExcelArgumentError);
  });
});

describe("setters check their values", () => {
  it.each([
    ["setHeaderRow", "yes"],
    ["setTotalsRow", 1],
    ["setTheme", 42],
    ["setShowFirstColumn", "x"],
    ["setShowLastColumn", 0],
    ["setShowRowStripes", 1],
    ["setShowColumnStripes", null]
  ] as const)("Table.%s refuses %s", (name, value) => {
    const { t } = setup();
    const before = Table.getModel(t);
    expect(() => (Table[name] as (t: unknown, v: unknown) => void)(t, value)).toThrow(
      ExcelArgumentError
    );
    expect(Table.getModel(t)).toEqual(before);
  });

  it.each([
    ["setColumnName", 42],
    ["setColumnFilterButton", "x"],
    ["setColumnTotalsRowLabel", 42],
    ["setColumnTotalsRowFormula", 42],
    ["setColumnTotalsRowResult", Number.NaN],
    ["setColumnTotalsRowResult", {}],
    ["setColumnTotalsRowFunction", "bogus"]
  ] as const)("Table.%s refuses %s", (name, value) => {
    const { t } = setup();
    const column = Table.getColumn(t, 0);
    expect(() => (Table[name] as (v: unknown, x: unknown) => void)(column, value)).toThrow(
      ExcelArgumentError
    );
  });

  it("Table.add checks every column's properties and the table's flags", () => {
    const { ws } = setup();
    const base = { name: "T_z", ref: "H1", rows: [[1]] };
    expect(() =>
      Table.add(ws, { ...base, columns: [{ name: "a", totalsRowLabel: 42 as never }] })
    ).toThrow(ExcelArgumentError);
    expect(() => Table.add(ws, { ...base, columns: [{ name: 42 as never }] })).toThrow(
      ExcelArgumentError
    );
    expect(() =>
      Table.add(ws, { ...base, headerRow: "yes" as never, columns: [{ name: "a" }] })
    ).toThrow(ExcelArgumentError);
    expect(() =>
      Table.add(ws, { ...base, style: { theme: 42 as never }, columns: [{ name: "a" }] })
    ).toThrow(ExcelArgumentError);
    expect(() => Table.add(ws, null as never)).toThrow(ExcelArgumentError);
    expect(() => Table.add(ws, { ...base, columns: [] })).toThrow(ExcelArgumentError);
    expect(Table.list(ws)).toHaveLength(1);
  });

  it("a written table carries what the setters were given", async () => {
    const { wb, t } = setup();
    Table.setColumnTotalsRowLabel(Table.getColumn(t, 0), "Sum");
    Table.setTotalsRow(t, true);
    await expect(Workbook.toBuffer(wb)).resolves.toBeInstanceOf(Uint8Array);
  });
});
