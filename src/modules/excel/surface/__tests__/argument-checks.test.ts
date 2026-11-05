/**
 * Arguments are checked where they are passed. Each of these used to be accepted and then stored as a value no
 * workbook can hold, dropped when the workbook was written, or written as something Excel repairs.
 */
import {
  Cell,
  Column,
  ExcelDefinedNameError,
  DefinedNames,
  ExcelArgumentError,
  Image,
  ExcelImageError,
  ExcelInvalidAddressError,
  ExcelInvalidValueTypeError,
  Row,
  ExcelRowOutOfBoundsError,
  ExcelColumnOutOfBoundsError,
  Table,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

const PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0,
  0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 15, 0, 0, 1, 1, 0, 5, 24,
  216, 78, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130
]);

describe("cell values a workbook cannot hold are refused", () => {
  it.each([NaN, Infinity, -Infinity])("%s", value => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", 1);
    expect(() => Cell.setValue(ws, "A1", value)).toThrow(ExcelInvalidValueTypeError);
    expect(Cell.getValue(ws, "A1")).toBe(1);
  });

  it("in every path that writes a value: rows, tables, formula results", () => {
    const { ws } = sheet();
    expect(() => Row.add(ws, [1, NaN])).toThrow(ExcelInvalidValueTypeError);
    expect(() =>
      Table.add(ws, { name: "T", ref: "D1", columns: [{ name: "A" }], rows: [[Infinity]] })
    ).toThrow(ExcelInvalidValueTypeError);
    expect(() => Cell.setValue(ws, "A5", { formula: "1/0", result: NaN })).toThrow(
      /formula result/
    );
    Cell.setValue(ws, "A6", { formula: "1+1", result: 2 });
    expect(() => Cell.setResult(ws, "A6", NaN)).toThrow(ExcelInvalidValueTypeError);
    expect(Cell.getResult(ws, "A6")).toBe(2);
  });

  it("functions, symbols, bigints and objects with no JSON text", () => {
    const { ws } = sheet();
    const cyclic: Record<string, unknown> = {};
    cyclic.self = cyclic;
    for (const value of [() => 1, Symbol("s"), 10n, cyclic, { big: 1n }]) {
      expect(() => Cell.setValue(ws, "A1", value as never)).toThrow(ExcelInvalidValueTypeError);
    }
    // A plain object is still the documented JSON value.
    Cell.setValue(ws, "A1", { foo: 1 } as never);
    expect(Cell.getText(ws, "A1")).toBe('{"foo":1}');
  });

  it("an error value is the way to write what NaN meant", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", { error: "#NUM!" });
    expect(Cell.getValue(ws, "A1")).toEqual({ error: "#NUM!" });
  });
});

describe("addresses, rows and columns", () => {
  it("a malformed address is an ExcelInvalidAddressError, for writers and readers alike", () => {
    const { ws } = sheet();
    for (const address of ["1A", "A0", "A99999999", "", "A1B"]) {
      expect(() => Cell.setValue(ws, address, 1), address).toThrow(ExcelInvalidAddressError);
      expect(() => Cell.getValue(ws, address), address).toThrow(ExcelInvalidAddressError);
    }
  });

  it("a row or column number outside the grid", () => {
    const { ws } = sheet();
    for (const row of [0, -1, 1.5, NaN, 1048577]) {
      expect(() => Cell.getValue(ws, row, 1), String(row)).toThrow(ExcelRowOutOfBoundsError);
      expect(() => Cell.setValue(ws, row, 1, 1), String(row)).toThrow(ExcelRowOutOfBoundsError);
      expect(() => Row.setHeight(ws, row, 10), String(row)).toThrow(ExcelRowOutOfBoundsError);
    }
    for (const col of [0, -1, 1.5, 16385]) {
      expect(() => Cell.setValue(ws, 1, col, 1), String(col)).toThrow(ExcelColumnOutOfBoundsError);
    }
    expect(() => Row.insert(ws, 0, [1])).toThrow(ExcelRowOutOfBoundsError);
    expect(Row.getCount(ws)).toBe(0);
  });
});

describe("row and column dimensions", () => {
  it("a height, width or outline level outside what Excel stores", () => {
    const { ws } = sheet();
    for (const height of [-5, NaN, Infinity, "10" as never]) {
      expect(() => Row.setHeight(ws, 1, height)).toThrow(ExcelArgumentError);
    }
    for (const width of [-1, NaN]) {
      expect(() => Column.setWidth(ws, 1, width)).toThrow(ExcelArgumentError);
    }
    for (const level of [-1, 8, 1.5]) {
      expect(() => Row.setOutlineLevel(ws, 1, level)).toThrow(ExcelArgumentError);
      expect(() => Column.setOutlineLevel(ws, 1, level)).toThrow(ExcelArgumentError);
    }
    expect(Row.getHeight(ws, 1)).toBeUndefined();
    Row.setHeight(ws, 1, 0);
    Row.setOutlineLevel(ws, 1, 7);
    Column.setWidth(ws, 1, 0);
    expect([Row.getHeight(ws, 1), Row.getOutlineLevel(ws, 1), Column.getWidth(ws, 1)]).toEqual([
      0, 7, 0
    ]);
  });
});

describe("defined names", () => {
  it("a name Excel would repair the file for, or a location that is not a reference", () => {
    const { wb } = sheet();
    const dn = Workbook.getDefinedNames(wb);
    for (const name of ["A1", "XFD1048576", "R1C1", "R", "c", "1st", "has space", ""]) {
      expect(() => DefinedNames.add(dn, "S!$A$1", name), name).toThrow(ExcelDefinedNameError);
      expect(() => DefinedNames.addFormula(dn, name, "1"), name).toThrow(ExcelDefinedNameError);
    }
    expect(() => DefinedNames.add(dn, "not a ref", "Good")).toThrow(ExcelDefinedNameError);
    expect(() => DefinedNames.add(dn, "S!$A$1:$A$99999999", "Good")).toThrow(ExcelDefinedNameError);
    expect(() => DefinedNames.addFormula(dn, "Good", "")).toThrow(ExcelDefinedNameError);
    expect(DefinedNames.list(dn)).toEqual([]);
  });

  it("names Excel accepts still work", () => {
    const { wb } = sheet();
    const dn = Workbook.getDefinedNames(wb);
    for (const name of ["Rate", "_hidden", "\\back", "Tax.2024", "XFE1", "ABCD1", "税率", "q?"]) {
      DefinedNames.add(dn, "S!$A$1", name);
    }
    DefinedNames.addFormula(dn, "Twice", "S!$A$1*2");
    expect(DefinedNames.getNames(dn, "S!A1")).toHaveLength(8);
  });
});

describe("images", () => {
  it("placing or using as background an image the workbook does not have", () => {
    const { wb, ws } = sheet();
    expect(() => Image.place(ws, 999, "A1:B2")).toThrow(ExcelImageError);
    expect(() => Image.setBackground(ws, 3)).toThrow(ExcelImageError);
    expect(Image.list(ws)).toEqual([]);
    const id = Image.add(wb, { buffer: PNG, extension: "png" });
    Image.place(ws, id, "A1:B2");
    expect(Image.list(ws)).toHaveLength(1);
  });
});

describe("a removed worksheet", () => {
  it("refuses edits and still answers reads", () => {
    const { wb, ws } = sheet();
    Cell.setValue(ws, "A1", 1);
    const t = Table.add(ws, { name: "T", ref: "C1", columns: [{ name: "A" }], rows: [[1]] });
    Workbook.removeWorksheet(wb, "S");
    expect(() => Cell.setValue(ws, "A1", 2)).toThrow(ExcelArgumentError);
    expect(() => Cell.setValue(ws, "Z9", 2)).toThrow(/was removed/);
    expect(() => Row.setHeight(ws, 1, 20)).toThrow(ExcelArgumentError);
    expect(() => Table.addRow(t, [2])).toThrow(ExcelArgumentError);
    expect(Cell.getValue(ws, "A1")).toBe(1);
    expect(Worksheet.getName(ws)).toBe("S");
  });
});

describe("style values outside the schema's vocabulary", () => {
  it.each<[string, (ws: Worksheet.Handle) => void]>([
    ["font size", ws => Cell.setFont(ws, "A1", { size: -3 })],
    ["font colour", ws => Cell.setFont(ws, "A1", { color: { argb: "red" } })],
    ["underline", ws => Cell.setFont(ws, "A1", { underline: "zig" as never })],
    ["border style", ws => Cell.setBorder(ws, "A1", { top: { style: "wavy" as never } })],
    ["fill pattern", ws => Cell.setFill(ws, "A1", { type: "pattern", pattern: "dots!" as never })],
    ["fill type", ws => Cell.setFill(ws, "A1", { type: "plaid" } as never)],
    ["alignment", ws => Cell.setAlignment(ws, "A1", { horizontal: "bogus" as never })],
    ["indent", ws => Cell.setAlignment(ws, "A1", { indent: -2 })],
    ["rotation", ws => Cell.setAlignment(ws, "A1", { textRotation: 500 })],
    ["number format", ws => Cell.setNumFmt(ws, "A1", 42 as never)],
    ["protection", ws => Cell.setProtection(ws, "A1", { locked: "yes" as never })],
    ["style", ws => Cell.setStyle(ws, "A1", { font: { size: 0 } })],
    ["row style", ws => Row.setFont(ws, 1, { size: -1 })],
    ["column style", ws => Column.setStyle(ws, 1, { border: { left: { style: "x" as never } } })],
    [
      "table column style",
      ws =>
        Table.add(ws, {
          name: "T",
          ref: "D1",
          columns: [{ name: "A", style: { font: { size: -1 } } }],
          rows: [[1]]
        })
    ]
  ])("%s", (_, apply) => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", 1);
    expect(() => apply(ws)).toThrow(ExcelArgumentError);
    expect(Cell.findStyle(ws, "A1")).toEqual({});
    expect(Table.list(ws)).toEqual([]);
  });

  it("refusing a style leaves no new cell behind", () => {
    const { ws } = sheet();
    expect(() => Cell.setFont(ws, "C9", { size: -1 })).toThrow(ExcelArgumentError);
    expect(() => Cell.setValue(ws, "C9", NaN)).toThrow(ExcelInvalidValueTypeError);
    expect(Cell.has(ws, "C9")).toBe(false);
  });

  it("every value Excel accepts still works", () => {
    const { ws } = sheet();
    Cell.setFont(ws, "A1", {
      name: "Arial",
      size: 409,
      bold: true,
      underline: "double",
      vertAlign: "superscript",
      color: { argb: "FF00FF00" },
      scheme: "minor"
    });
    Cell.setFont(ws, "A2", { color: { argb: "00FF00" } });
    Cell.setFont(ws, "A3", { color: { theme: 1, tint: -0.5 } });
    Cell.setBorder(ws, "A1", { diagonal: { style: "thin", up: true } });
    Cell.setFill(ws, "A1", {
      type: "gradient",
      gradient: "angle",
      degree: 90,
      stops: [
        { position: 0, color: { argb: "FFFFFFFF" } },
        { position: 1, color: { argb: "FF000000" } }
      ]
    });
    Cell.setAlignment(ws, "A1", { textRotation: "vertical", indent: 3, horizontal: "center" });
    Cell.setAlignment(ws, "A2", { textRotation: 255 });
    Cell.setNumFmt(ws, "A1", "0.00");
    Cell.setFont(ws, "A1", undefined);
    expect(Cell.getNumFmt(ws, "A1")).toBe("0.00");
  });
});

describe("row values", () => {
  it("a row is a list, an object keyed by column key, or nothing", () => {
    const { ws } = sheet();
    for (const bad of [5, "text", true]) {
      expect(() => Row.add(ws, bad as never)).toThrow(ExcelArgumentError);
      expect(() => Row.insert(ws, 1, bad as never)).toThrow(ExcelArgumentError);
      expect(() => Row.setValues(ws, 1, bad as never)).toThrow(ExcelArgumentError);
      expect(() => Row.addMany(ws, [[1], bad as never])).toThrow(/values\[1\]/);
    }
    expect(Row.getCount(ws)).toBe(0);
    Row.add(ws, [1, 2]);
    Row.add(ws, null);
    expect(Row.getCount(ws)).toBe(2);
  });
});
