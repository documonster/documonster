/**
 * The implicit-intersection operator `@` as Excel 16.114 writes it in XLSB, from two workbooks this library wrote as
 * XLSX (`@x` stored as `_xlfn.SINGLE(x)`) and Excel saved as XLSB (`tmp/office-tests/r5-single/README.md`):
 *
 * - `single.excel.xlsb`: `@` in a legacy array formula, a dynamic array, a conditional format, a data validation and
 *   a defined name — and two plain cells.
 * - `single-cells.excel.xlsb`: `@` in eighteen cell-formula positions.
 *
 * What they show: `@x` is a call to the `_xlfn.SINGLE` stub — `PtgName`, the operand, `PtgFuncVar(255)` with the class
 * of the call's position — everywhere except in a cell formula where `x` alone would be intersected and displayed
 * with `@` by Excel anyway (`@A1:A3+1`, `@Rng`, `@SEQUENCE(3)`). There Excel writes `x` bare, in the value class.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, ConditionalFormatting, DataValidation, DefinedNames, Workbook } from "@excel";
import { encodeCol } from "@excel/utils/address";
import { XlsbReader, iterateInterpretableRecords, readCell } from "@excel/xlsb/binary";
import { encodeFormula, formulaText, type PtgContext } from "@excel/xlsb/formula/ptg";
import { readWorkbookPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";

function hex(bytes: Uint8Array): string {
  return [...bytes].map(byte => byte.toString(16).padStart(2, "0")).join(" ");
}

async function excelWorkbook(name: string) {
  const archive = await extractAll(await readFile(`${DIRECTORY}/${name}`));
  const workbookPart = archive.get("xl/workbook.bin")!.data;
  const { formulaContext } = readWorkbookPart(workbookPart, "xl/workbook.bin");
  return {
    workbookPart,
    sheet: archive.get("xl/worksheets/sheet1.bin")!.data,
    context: { ...formulaContext, sheetIndex: 0 } as PtgContext
  };
}

interface FormulaRecord {
  readonly address: string;
  readonly rgce: Uint8Array;
  readonly extra: Uint8Array;
}

/** Each `BrtFmla*` cell's tokens, and each `BrtArrFmla`'s, by the address of its first cell. */
function formulaRecords(sheet: Uint8Array): { cells: FormulaRecord[]; arrays: FormulaRecord[] } {
  const cells: FormulaRecord[] = [];
  const arrays: FormulaRecord[] = [];
  let row = 0;
  for (const record of iterateInterpretableRecords(sheet, "sheet")) {
    const name = recordSpec(record.id)?.name;
    const reader = new XlsbReader(record.payload, 0, "sheet");
    if (name === "BrtRowHdr") {
      row = reader.readUint32();
      continue;
    }
    if (name === "BrtArrFmla") {
      // `RfX` (four 32-bit bounds), the flags byte, then the formula.
      const first = reader.readUint32();
      reader.skip(4);
      const column = reader.readUint32();
      reader.skip(4 + 1);
      const rgce = reader.readBytes(reader.readUint32());
      const extra = reader.readBytes(reader.readUint32());
      arrays.push({ address: `${encodeCol(column)}${first + 1}`, rgce, extra });
      continue;
    }
    if (!name?.startsWith("BrtFmla")) {
      continue;
    }
    const { column } = readCell(reader);
    if (name === "BrtFmlaNum") {
      reader.readFloat64();
    } else if (name === "BrtFmlaString") {
      reader.skip(reader.readUint32() * 2);
    } else {
      reader.readUint8();
    }
    reader.readUint16();
    const rgce = reader.readBytes(reader.readUint32());
    const extra = reader.readBytes(reader.readUint32());
    if (rgce[0] !== 0x01) {
      cells.push({ address: `${encodeCol(column)}${row + 1}`, rgce, extra });
    }
  }
  return { cells, arrays };
}

/** Every `BrtName` record of a workbook part, whole, by the name it defines. */
function nameRecords(workbookPart: Uint8Array): Map<string, Uint8Array> {
  const records = new Map<string, Uint8Array>();
  for (const record of iterateInterpretableRecords(workbookPart, "workbook")) {
    if (recordSpec(record.id)?.name !== "BrtName") {
      continue;
    }
    const reader = new XlsbReader(record.payload, 0, "workbook");
    reader.skip(9);
    const name = String.fromCharCode(
      ...Array.from({ length: reader.readUint32() }, () => reader.readUint16())
    );
    records.set(name, record.payload);
  }
  return records;
}

/** The record's bytes contain `needle`. */
function contains(haystack: Uint8Array, needle: string): boolean {
  return hex(haystack).includes(needle);
}

/**
 * `single-cells`: each cell's text as the library reads the XLSX it wrote, and what Excel's XLSB of it reads back as.
 * Where Excel dropped the call the bare operand reads back without `@` — the same formula, as the XLSX `<f>A1:A3</f>`
 * reads `A1:A3`; where Excel kept the call, `@` reads back.
 */
const CELLS: readonly (readonly [address: string, text: string, read: string])[] = [
  ["C1", "@A1:A3+1", "A1:A3+1"],
  ["C2", "ABS(@A1:A3)", "ABS(A1:A3)"],
  ["C3", '@A1:A3&"x"', 'A1:A3&"x"'],
  ["C4", "IF(@A1:A3>1,1,0)", "IF(A1:A3>1,1,0)"],
  ["C5", "INDEX(@A1:A3,1)", "INDEX(@A1:A3,1)"],
  ["C6", "ROWS(@A1:A3)", "ROWS(@A1:A3)"],
  ["C7", "SUMPRODUCT(@A1:A3)", "SUMPRODUCT(@A1:A3)"],
  ["C8", "@A1", "@A1"],
  ["C10", "-@A1:A3", "-A1:A3"],
  ["C11", "@INDEX(A1:A3,0)", "INDEX(A1:A3,0)"],
  ["C12", "@_xlfn.SEQUENCE(3)", "_xlfn.SEQUENCE(3)"],
  ["C13", "SUM(@A1)", "SUM(@A1)"],
  ["C14", "SUM(@_xlfn.SEQUENCE(3))", "SUM(@_xlfn.SEQUENCE(3))"],
  ["C15", "@Rng", "Rng"],
  ["C16", "SUM(@Rng)", "SUM(@Rng)"],
  ["C17", "SUM((@A1:A3))", "SUM((@A1:A3))"],
  ["C18", "@'S'!A1:A3", "S!A1:A3"]
];

describe("@ in a cell formula: the tokens Excel wrote", () => {
  it("are what this encoder writes for the same text, and read back to the same formula", async () => {
    const { sheet, context } = await excelWorkbook("single-cells.excel.xlsb");
    const source = Workbook.getWorksheet(
      await Workbook.read(await readFile(`${DIRECTORY}/single-cells.xlsx`)),
      "S"
    )!;
    const records = new Map(formulaRecords(sheet).cells.map(cell => [cell.address, cell]));
    const differing: string[] = [];
    for (const [address, text, read] of CELLS) {
      expect(Cell.getFormula(source, address)).toBe(text);
      const excel = records.get(address)!;
      const ours = encodeFormula(text, context, address);
      if (hex(ours.rgce) !== hex(excel.rgce)) {
        differing.push(`${address} ${text}\n  ours  ${hex(ours.rgce)}\n  excel ${hex(excel.rgce)}`);
      }
      expect(formulaText(excel.rgce, context, address, excel.extra)).toBe(read);
      expect(formulaText(ours.rgce, context, address, ours.extra)).toBe(read);
    }
    expect(differing).toEqual([]);
  });

  it("pins the encodings: the call, with the class of its position, or the operand bare", async () => {
    const { sheet } = await excelWorkbook("single-cells.excel.xlsb");
    const records = new Map(
      formulaRecords(sheet).cells.map(cell => [cell.address, hex(cell.rgce)])
    );
    // `_xlfn.SINGLE` is defined name 2 here (`_xlfn.SEQUENCE` is 1, `Rng` 3).
    expect(records.get("C5")).toBe(
      "23 02 00 00 00 25 00 00 00 00 02 00 00 00 00 c0 00 c0 22 02 ff 00 1e 01 00 42 02 1d 00"
    );
    expect(records.get("C8")).toBe("23 02 00 00 00 24 00 00 00 00 00 c0 42 02 ff 00");
    expect(records.get("C1")).toBe("45 00 00 00 00 02 00 00 00 00 c0 00 c0 1e 01 00 03");
    expect(records.get("C15")).toBe("43 03 00 00 00");
  });

  /**
   * `@A1:A3:B1`: Excel writes the reference expression as a value-class `PtgMemArea` carrying the area it computed
   * (`46 …`, with the cache in `RgbExtra`), where this encoder writes the `PtgMemFunc` it writes for every reference
   * expression (`49 …`) — no cache to keep current. Both mean the same and read back alike; the `@` is dropped by both.
   */
  it("@A1:A3:B1 is the intersected reference expression in both, Excel's carrying a cache", async () => {
    const { sheet, context } = await excelWorkbook("single-cells.excel.xlsb");
    const excel = formulaRecords(sheet).cells.find(cell => cell.address === "C9")!;
    expect(hex(excel.rgce).startsWith("46 ")).toBe(true);
    const ours = encodeFormula("@A1:A3:B1", context, "C9");
    expect(hex(ours.rgce).startsWith("49 ")).toBe(true);
    expect(formulaText(excel.rgce, context, "C9", excel.extra)).toBe("A1:A3:B1");
    expect(formulaText(ours.rgce, context, "C9", ours.extra)).toBe("A1:A3:B1");
  });
});

describe("@ outside a cell formula: the tokens Excel wrote", () => {
  // `_xlfn.SINGLE` is defined name 1 in `single.excel.xlsb`.
  const SINGLE_AREA = "23 01 00 00 00 25 00 00 00 00 02 00 00 00 00 00 00 00";

  it("a legacy array formula and a dynamic array call _xlfn.SINGLE", async () => {
    const { sheet, context } = await excelWorkbook("single.excel.xlsb");
    const arrays = formulaRecords(sheet).arrays;
    expect(arrays.map(array => array.address)).toEqual(["E1", "G1"]);
    const expected = `45 00 00 00 00 02 00 00 00 00 c0 00 c0 23 01 00 00 00 25 00 00 00 00 02 00 00 00 00 c0 00 c0 42 02 ff 00 05`;
    const ours = encodeFormula("A1:A3*@A1:A3", { ...context, role: "array" }, "E1");
    for (const array of arrays) {
      expect(hex(array.rgce)).toBe(expected);
      expect(
        formulaText(array.rgce, { ...context, role: "array" }, array.address, array.extra)
      ).toBe("A1:A3*@A1:A3");
    }
    expect(hex(ours.rgce)).toBe(expected);
  });

  it.each([
    ["conditional", "@$A$1:$A$3>1", `${SINGLE_AREA} 42 02 ff 00 1e 01 00 0d`],
    ["validation", "@$A$1:$A$3>0", `${SINGLE_AREA} 42 02 ff 00 1e 00 00 0d`]
  ] as const)("a %s formula calls _xlfn.SINGLE", async (role, text, expected) => {
    const { sheet, context } = await excelWorkbook("single.excel.xlsb");
    const record = [...iterateInterpretableRecords(sheet, "sheet")].find(each =>
      contains(each.payload, expected)
    );
    expect(recordSpec(record!.id)?.name).toBe(
      role === "conditional" ? "BrtBeginCFRule" : "BrtDVal"
    );
    const ours = encodeFormula(text, { ...context, role }, "B1");
    expect(hex(ours.rgce)).toBe(expected);
    const tokens = Uint8Array.from(expected.split(" ").map(byte => parseInt(byte, 16)));
    expect(formulaText(tokens, { ...context, role }, "B1")).toBe(text);
  });

  it("a defined name calls _xlfn.SINGLE in the reference class", async () => {
    const { workbookPart, context } = await excelWorkbook("single.excel.xlsb");
    const expected = "23 01 00 00 00 3b 00 00 00 00 00 00 02 00 00 00 00 00 00 00 22 02 ff 00";
    expect(contains(nameRecords(workbookPart).get("N")!, expected)).toBe(true);
    const ours = encodeFormula(
      "@S!$A$1:$A$3",
      { ...context, role: "name", sheetIndex: undefined },
      "N"
    );
    expect(hex(ours.rgce)).toBe(expected);
    const tokens = Uint8Array.from(expected.split(" ").map(byte => parseInt(byte, 16)));
    expect(formulaText(tokens, { ...context, role: "name" }, "N")).toBe("@S!$A$1:$A$3");
  });

  /**
   * All but `fCalcExp` (bit 4 of the first byte). Excel sets that bit on the `_xlfn.SINGLE` stub of
   * `future-functions`, `formula-classes` and `let-oracle` (`1b 00 02 00`) and not on that of `formula-edges`,
   * `formula-matrix`, `single` or `single-cells` (`0b 00 02 00`) — whether formulas call it or not, in whichever
   * context — and likewise on some workbooks' `_xlfn.SORTBY` and `_xlfn.LAMBDA` stubs and not others'. No rule in
   * these files decides it, and Excel opens either; this writer keeps setting it (`CALC_EXP_FUTURE_FUNCTIONS`).
   */
  it("the _xlfn.SINGLE stub this writer emits is Excel's record, but for fCalcExp", async () => {
    const { workbookPart } = await excelWorkbook("single.excel.xlsb");
    const workbook = await Workbook.read(await readFile(`${DIRECTORY}/single.xlsx`));
    const ours = nameRecords(
      (await extractAll(await Workbook.toBuffer(workbook, { format: "xlsb" }))).get(
        "xl/workbook.bin"
      )!.data
    ).get("_xlfn.SINGLE")!;
    const excel = nameRecords(workbookPart).get("_xlfn.SINGLE")!;
    expect(ours[0]).toBe(excel[0]! | 0x10);
    expect(hex(ours.subarray(1))).toBe(hex(excel.subarray(1)));
  });
});

describe("@ through this library's XLSB", () => {
  it("reads back as written, in every context", async () => {
    const workbook = await Workbook.read(await readFile(`${DIRECTORY}/single.xlsx`));
    const read = await Workbook.read(await Workbook.toBuffer(workbook, { format: "xlsb" }));
    const sheet = Workbook.getWorksheet(read, "S")!;
    expect(Cell.getFormula(sheet, "E1")).toBe("A1:A3*@A1:A3");
    expect(Cell.getFormula(sheet, "G1")).toBe("A1:A3*@A1:A3");
    expect(Cell.getFormula(sheet, "C5")).toBe("SUM(@A1:A3)");
    expect(ConditionalFormatting.list(sheet)[0]!.rules[0]).toMatchObject({
      formulae: ["@$A$1:$A$3>1"]
    });
    expect(DataValidation.find(sheet, "D1")).toMatchObject({ formulae: ["@$A$1:$A$3>0"] });
    expect(DefinedNames.get(Workbook.getDefinedNames(read), "N")?.formulaExpression).toBe(
      "@S!$A$1:$A$3"
    );
  });

  // Byte for byte against Excel's name table in the first test; this writer numbers its names in its own order.
  it("writes the cells of single-cells to tokens that read back as Excel's do", async () => {
    const workbook = await Workbook.read(await readFile(`${DIRECTORY}/single-cells.xlsx`));
    const ours = await extractAll(await Workbook.toBuffer(workbook, { format: "xlsb" }));
    const { formulaContext: context } = readWorkbookPart(
      ours.get("xl/workbook.bin")!.data,
      "xl/workbook.bin"
    );
    const cells = new Map(
      formulaRecords(ours.get("xl/worksheets/sheet1.bin")!.data).cells.map(cell => [
        cell.address,
        cell
      ])
    );
    for (const [address, , read] of CELLS) {
      const cell = cells.get(address)!;
      expect(formulaText(cell.rgce, { ...context, sheetIndex: 0 }, address, cell.extra)).toBe(read);
    }
  });
});
