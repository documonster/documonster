/**
 * A bare function name as an argument — an eta-reduced lambda, `GROUPBY(A2:A9,B2:B9,SUM)` — against Excel's own XLSB.
 *
 * `eta-lambda.excel.xlsb` is Excel 16 (macOS) typing eight such formulas into a sheet and saving it; the `.xlsx` is
 * the same workbook saved by Excel as XLSX, which spells each name `_xleta.SUM`. In XLSB the name is a `PtgName` to a
 * hidden `BrtName` called `_xleta.SUM` — flags `0x00020019`, body `#NAME?` — whatever the function: a future one
 * (`PERCENTOF`, `ARRAYTOTEXT`) gets no `_xlfn.` inside the `_xleta.` name.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, DefinedNames, Workbook } from "@excel";
import { encodeCol } from "@excel/utils/address";
import { XlsbReader, iterateInterpretableRecords, readCell } from "@excel/xlsb/binary";
import { encodeFormula, formulaText } from "@excel/xlsb/formula/ptg";
import { readWorkbookPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";

/** Excel's XLSX text of each formula cell, and the same formula as typed. */
const FORMULAS: readonly (readonly [string, string, string])[] = [
  ["D2", "_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.SUM)", "GROUPBY(A2:A9,B2:B9,SUM)"],
  ["G2", "_xlfn.PIVOTBY(A2:A9,A2:A9,B2:B9,_xleta.SUM)", "PIVOTBY(A2:A9,A2:A9,B2:B9,SUM)"],
  ["D8", "_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.AVERAGE,0,0)", "GROUPBY(A2:A9,B2:B9,AVERAGE,0,0)"],
  ["D12", "_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.PERCENTOF,0,0)", "GROUPBY(A2:A9,B2:B9,PERCENTOF,0,0)"],
  ["D16", "_xlfn.BYROW(B2:B4,_xleta.SUM)", "BYROW(B2:B4,SUM)"],
  ["F16", "_xlfn.MAP(B2:B4,_xleta.ABS)", "MAP(B2:B4,ABS)"],
  [
    "H16",
    "SUM(B2:B4)+_xlfn.GROUPBY(A2:A3,B2:B3,_xleta.COUNTA,0,0)",
    "SUM(B2:B4)+GROUPBY(A2:A3,B2:B3,COUNTA,0,0)"
  ],
  ["D20", "_xlfn.BYCOL(B2:B4,_xleta.ARRAYTOTEXT)", "BYCOL(B2:B4,ARRAYTOTEXT)"]
];

async function parts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await extractAll(bytes);
  return new Map([...archive].map(([path, entry]) => [path, entry.data]));
}

/** Each formula cell's token stream, by address. */
function formulaTokens(sheet: Uint8Array): Map<string, { rgce: Uint8Array; extra: Uint8Array }> {
  const cells = new Map<string, { rgce: Uint8Array; extra: Uint8Array }>();
  let row = 0;
  for (const record of iterateInterpretableRecords(sheet, "sheet")) {
    const name = recordSpec(record.id)?.name;
    if (name === "BrtRowHdr") {
      row = new DataView(record.payload.buffer, record.payload.byteOffset).getUint32(0, true);
      continue;
    }
    if (name !== "BrtFmlaNum" && name !== "BrtFmlaString" && name !== "BrtFmlaBool") {
      continue;
    }
    const reader = new XlsbReader(record.payload, 0, "sheet");
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
    cells.set(`${encodeCol(column)}${row + 1}`, { rgce, extra });
  }
  return cells;
}

/** The dynamic-array formulas live in `BrtArrFmla` records; each master's tokens, by address. */
function arrayFormulaTokens(sheet: Uint8Array): Map<string, string> {
  const cells = new Map<string, string>();
  for (const record of iterateInterpretableRecords(sheet, "sheet")) {
    if (recordSpec(record.id)?.name !== "BrtArrFmla") {
      continue;
    }
    const reader = new XlsbReader(record.payload, 0, "sheet");
    const firstRow = reader.readUint32();
    reader.readUint32();
    const firstColumn = reader.readUint32();
    reader.readUint32();
    reader.readUint8();
    const rgce = reader.readBytes(reader.readUint32());
    cells.set(`${encodeCol(firstColumn)}${firstRow + 1}`, hex(rgce));
  }
  return cells;
}

function hex(bytes: Uint8Array): string {
  return [...bytes].map(byte => byte.toString(16).padStart(2, "0")).join(" ");
}

/** Every `BrtName` payload whose name starts `_xleta.`, by name. */
function etaNames(workbookPart: Uint8Array): Map<string, string> {
  const records = new Map<string, string>();
  for (const record of iterateInterpretableRecords(workbookPart, "workbook")) {
    if (recordSpec(record.id)?.name !== "BrtName") {
      continue;
    }
    const reader = new XlsbReader(record.payload, 0, "workbook");
    reader.skip(9);
    const name = String.fromCharCode(
      ...Array.from({ length: reader.readUint32() }, () => reader.readUint16())
    );
    if (name.startsWith("_xleta.")) {
      records.set(name, hex(record.payload));
    }
  }
  return records;
}

describe("eta-reduced lambdas: Excel's XLSB", () => {
  it("reads to the text Excel's XLSX holds, with no `_xleta.` name listed", async () => {
    const workbook = await Workbook.read(
      Workbook.create(),
      await readFile(`${DIRECTORY}/eta-lambda.excel.xlsb`)
    );
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    expect(FORMULAS.map(([address]) => Cell.getFormula(sheet, address))).toEqual(
      FORMULAS.map(([, text]) => text)
    );
    expect(DefinedNames.list(Workbook.getDefinedNames(workbook))).toEqual([]);
  });

  it("encodes each formula, stored or typed, to Excel's tokens", async () => {
    const files = await parts(await readFile(`${DIRECTORY}/eta-lambda.excel.xlsb`));
    const { formulaContext } = readWorkbookPart(files.get("xl/workbook.bin")!, "xl/workbook.bin");
    const excel = arrayFormulaTokens(files.get("xl/worksheets/sheet1.bin")!);
    const context = { ...formulaContext, sheetIndex: 0 };
    for (const [address, stored, typed] of FORMULAS) {
      expect(excel.has(address)).toBe(true);
      expect(hex(encodeFormula(stored, context, address).rgce), stored).toBe(excel.get(address));
      expect(hex(encodeFormula(typed, context, address).rgce), typed).toBe(excel.get(address));
      expect(formulaText(encodeFormula(typed, context, address).rgce, context, address)).toBe(
        stored
      );
    }
    // Every cell of a spill is a `BrtFmla*` too, carrying no tokens of its own.
    expect([...formulaTokens(files.get("xl/worksheets/sheet1.bin")!).values()]).not.toEqual([]);
  });

  it.each(["stored", "typed"] as const)(
    "writes Excel's `_xleta.` name records for %s formulas",
    async spelling => {
      const excel = etaNames(
        (await parts(await readFile(`${DIRECTORY}/eta-lambda.excel.xlsb`))).get("xl/workbook.bin")!
      );
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "S");
      for (let row = 2; row <= 9; row++) {
        Cell.setValue(sheet, `A${row}`, row % 2 ? "x" : "y");
        Cell.setValue(sheet, `B${row}`, row);
      }
      for (const [address, stored, typed] of FORMULAS) {
        Cell.setValue(sheet, address, { formula: spelling === "stored" ? stored : typed });
      }
      const ours = etaNames(
        (await parts(await Workbook.toBuffer(workbook, { format: "xlsb" }))).get("xl/workbook.bin")!
      );
      expect([...ours.keys()].sort()).toEqual([...excel.keys()].sort());
      for (const [name, bytes] of ours) {
        expect(bytes, name).toBe(excel.get(name));
      }
    }
  );

  it("does not read a function name as a function where none can be passed", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Cell.setValue(sheet, "A1", { formula: "COS(NA)", result: { error: "#NAME?" } });
    // `NA` here is a name the workbook does not define, which XLSB cannot encode — not `_xleta.NA`.
    await expect(Workbook.toBuffer(workbook, { format: "xlsb" })).rejects.toThrow(/the name NA/);
  });
});
