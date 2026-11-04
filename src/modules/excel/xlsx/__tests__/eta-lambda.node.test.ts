/**
 * A function passed by name — `GROUPBY(A2:A9,B2:B9,SUM)` — written to XLSX as Excel writes it: `_xleta.SUM`.
 *
 * Without the prefix Excel reads `SUM` as a defined name it does not have; its XLSB of such a file shows it creating an
 * empty name `SUM`, and the cell computes `#NAME?`. The oracle is Excel's own XLSX of the same formulas
 * (`xlsb/__tests__/data/eta-lambda.excel.xlsx`). A bare name the workbook *does* define stays that name.
 *
 * `.node` because it reads the fixture off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  DefinedNames,
  Stream,
  Table,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

const FIXTURE = "src/modules/excel/xlsb/__tests__/data/eta-lambda.excel.xlsx";

/** Typed formulas, by the address Excel's sample holds them at. */
const TYPED: readonly (readonly [string, string])[] = [
  ["D2", "GROUPBY(A2:A9,B2:B9,SUM)"],
  ["G2", "PIVOTBY(A2:A9,A2:A9,B2:B9,SUM)"],
  ["D8", "GROUPBY(A2:A9,B2:B9,AVERAGE,0,0)"],
  ["D12", "GROUPBY(A2:A9,B2:B9,PERCENTOF,0,0)"],
  ["D16", "BYROW(B2:B4,SUM)"],
  ["F16", "MAP(B2:B4,ABS)"],
  ["H16", "SUM(B2:B4)+GROUPBY(A2:A3,B2:B3,COUNTA,0,0)"],
  ["D20", "BYCOL(B2:B4,ARRAYTOTEXT)"]
];

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode((await extractAll(bytes)).get(path)!.data);
}

/** Each `<f>` text of a sheet part, by cell address. */
function cellFormulas(xml: string): Map<string, string> {
  return new Map(
    [...xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*>\s*<f[^>]*>([^<]*)<\/f>/g)].map(m => [m[1]!, m[2]!])
  );
}

describe("XLSX: functions passed by name", () => {
  it("writes each typed formula as Excel's XLSX holds it", async () => {
    const excel = cellFormulas(await part(await readFile(FIXTURE), "xl/worksheets/sheet1.xml"));
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    for (const [address, formula] of TYPED) {
      Cell.setValue(sheet, address, { formula });
    }
    const ours = cellFormulas(
      await part(await Workbook.toBuffer(workbook), "xl/worksheets/sheet1.xml")
    );
    for (const [address] of TYPED) {
      expect(ours.get(address), address).toBe(excel.get(address));
    }
  });

  it("reads Excel's file back in the stored form, and writes that form unchanged", async () => {
    const workbook = await Workbook.read(Workbook.create(), await readFile(FIXTURE));
    const sheet = Workbook.getWorksheet(workbook, "S")!;
    expect(Cell.getFormula(sheet, "D2")).toBe("_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.SUM)");
    const again = cellFormulas(
      await part(await Workbook.toBuffer(workbook), "xl/worksheets/sheet1.xml")
    );
    expect(again.get("D2")).toBe("_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.SUM)");
  });

  it("leaves a name the workbook defines as that name", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "ABS", "LAMBDA(x,x*2)");
    Cell.setValue(sheet, "A1", { formula: "MAP(B2:B4,ABS)" });
    Cell.setValue(sheet, "A2", { formula: "MAP(B2:B4,SUM)" });
    const xml = cellFormulas(
      await part(await Workbook.toBuffer(workbook), "xl/worksheets/sheet1.xml")
    );
    expect(xml.get("A1")).toBe("_xlfn.MAP(B2:B4,ABS)");
    expect(xml.get("A2")).toBe("_xlfn.MAP(B2:B4,_xleta.SUM)");
  });

  it("covers shared formulas, validations, conditional formats, names and table columns", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Cell.setValue(sheet, "A1", { formula: "BYROW(B1:C1,SUM)" });
    Cell.setValue(sheet, "A2", { sharedFormula: "A1", result: 0 } as never);
    DataValidation.add(sheet.dataValidations, "E1", {
      type: "custom",
      formulae: ["SUM(BYROW(B1:C1,SUM))>0"]
    });
    ConditionalFormatting.add(sheet, {
      ref: "F1",
      rules: [
        { type: "expression", formulae: ["SUM(MAP(B1:C1,ABS))>0"], style: { font: { bold: true } } }
      ]
    });
    DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Rows", "BYROW(S!$B$1:$C$2,SUM)");
    Table.add(sheet, {
      name: "T",
      ref: "H1",
      headerRow: true,
      columns: [{ name: "a" }, { name: "b", calculatedColumnFormula: "SUM(MAP(T[@a],ABS))" }],
      rows: [[1, { formula: "SUM(MAP(T[@a],ABS))", result: 1 }]]
    });
    const bytes = await Workbook.toBuffer(workbook);
    const sheetXml = await part(bytes, "xl/worksheets/sheet1.xml");
    expect(sheetXml).toContain("_xlfn.BYROW(B1:C1,_xleta.SUM)");
    expect(sheetXml).toContain("SUM(_xlfn.BYROW(B1:C1,_xleta.SUM))&gt;0");
    expect(sheetXml).toContain("SUM(_xlfn.MAP(B1:C1,_xleta.ABS))&gt;0");
    expect(await part(bytes, "xl/workbook.xml")).toContain("_xlfn.BYROW(S!$B$1:$C$2,_xleta.SUM)");
    expect(await part(bytes, "xl/tables/table1.xml")).toContain(
      "_xlfn.MAP(T[[#This Row],[a]],_xleta.ABS)"
    );
  });

  it("covers the streaming writer", async () => {
    const chunks: Uint8Array[] = [];
    const sink = new WritableStream<Uint8Array>({
      write(chunk) {
        chunks.push(chunk);
      }
    });
    const writer = new Stream.WorkbookWriter({ stream: sink as never });
    const ws = writer.addWorksheet("S");
    Stream.setCellValue(ws.getCell("A1"), { formula: "MAP(B2:B4,ABS)" });
    ws.commit();
    await writer.commit();
    const bytes = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
    let at = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, at);
      at += chunk.length;
    }
    // `MAP` produces an array, so it is a dynamic array — as the buffered writer stores it.
    expect(await part(bytes, "xl/worksheets/sheet1.xml")).toContain(
      '<f t="array" ref="A1">_xlfn.MAP(B2:B4,_xleta.ABS)</f>'
    );
  });
});
