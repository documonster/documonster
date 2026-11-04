/**
 * Future functions are written with the prefix Excel needs in the file.
 *
 * `FILTER(…)` written verbatim into `<f>` makes Excel refuse the whole workbook, and `XLOOKUP(…)`, `SEQUENCE(…)`,
 * `UNIQUE(…)`, `VSTACK(…)` and `LET(…)` open as `#NAME?`: unprefixed, they are user-defined functions to Excel.
 * Every place a formula is written gets the prefix — cells (plain, shared, array), conditional formats, data
 * validations, defined names, table columns, and the streaming writer, which renders cells through the same xform.
 */
import { extractAll } from "@archive/unzip/extract";
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  DefinedNames,
  Stream,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

const decoder = new TextDecoder();

async function parts(bytes: Uint8Array): Promise<{ sheet: string; workbook: string }> {
  const archive = await extractAll(bytes);
  return {
    sheet: decoder.decode(archive.get("xl/worksheets/sheet1.xml")!.data),
    workbook: decoder.decode(archive.get("xl/workbook.xml")!.data)
  };
}

describe("future-function prefixes in XLSX", () => {
  it("writes them in cells, shared and array formulas, CF, DV and defined names", async () => {
    const handle = Workbook.create();
    const ws = Workbook.addWorksheet(handle, "S");
    Cell.setValue(ws, "A1", { formula: "SUM(FILTER(B1:B4,B1:B4>1))" });
    Cell.setValue(ws, "A2", { formula: "LET(x,2,x*3)" });
    // A shared formula whose result is one value: one that produces an array is a dynamic array, which is not shared.
    Cell.setValue(ws, "A3", { formula: "XMATCH(2,B1:B4)" });
    Cell.setValue(ws, "A4", { sharedFormula: "A3" });
    Cell.setValue(ws, "A5", { formula: "SORT(B1:B4)", shareType: "array", ref: "A5:A8" });
    ConditionalFormatting.add(ws, {
      ref: "C1:C4",
      rules: [{ type: "expression", formulae: ["COUNTA(UNIQUE($B$1:$B$4))>1"], priority: 1 }]
    });
    DataValidation.add(ws.dataValidations, "D1", {
      type: "custom",
      formulae: ["ROWS(VSTACK(B1,B2))=2"]
    });
    DefinedNames.addFormula(Workbook.getDefinedNames(handle), "Seq", "SEQUENCE(3)");
    DefinedNames.addFormula(Workbook.getDefinedNames(handle), "Add", "LAMBDA(x,y,x+y)");
    const { sheet, workbook } = await parts(await Workbook.toBuffer(handle));
    expect(sheet).toContain("<f>SUM(_xlfn._xlws.FILTER(B1:B4,B1:B4&gt;1))</f>");
    expect(sheet).toContain("<f>_xlfn.LET(_xlpm.x,2,_xlpm.x*3)</f>");
    expect(sheet).toContain('ref="A3:A4" si="0">_xlfn.XMATCH(2,B1:B4)</f>');
    expect(sheet).toContain('<f t="array" ref="A5:A8">_xlfn._xlws.SORT(B1:B4)</f>');
    expect(sheet).toContain("<formula>COUNTA(_xlfn.UNIQUE($B$1:$B$4))&gt;1</formula>");
    expect(sheet).toContain("<formula1>ROWS(_xlfn.VSTACK(B1,B2))=2</formula1>");
    expect(workbook).toContain('<definedName name="Seq">_xlfn.SEQUENCE(3)</definedName>');
    expect(workbook).toContain(
      '<definedName name="Add">_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x+_xlpm.y)</definedName>'
    );
  });

  it("writes a this-row reference as [#This Row], and reads it back in that form", async () => {
    // `@` is the formula bar's spelling. Excel refuses a file holding `Tb[@[x]:[y]]` and reads `Tb[@x]:Tb[@y]` as
    // `#REF!:#REF!`; what it writes itself, and what a reader of its files has always returned, is `[#This Row]`.
    const handle = Workbook.create();
    const ws = Workbook.addWorksheet(handle, "S");
    Cell.setValue(ws, "A2", { formula: "SUM(Tb[@[x]:[y]])" });
    Cell.setValue(ws, "A3", { formula: "SUM(Tb[@x]:Tb[@y])" });
    DataValidation.add(ws.dataValidations, "D1", { type: "custom", formulae: ["Tb[@x]>0"] });
    const bytes = await Workbook.toBuffer(handle);
    const { sheet } = await parts(bytes);
    expect(sheet).toContain("<f>SUM(Tb[[#This Row],[x]:[y]])</f>");
    expect(sheet).toContain("<f>SUM(Tb[[#This Row],[x]]:Tb[[#This Row],[y]])</f>");
    expect(sheet).toContain("<formula1>Tb[[#This Row],[x]]&gt;0</formula1>");
    expect(Cell.getFormula(ws, "A2")).toBe("SUM(Tb[@[x]:[y]])");
    const back = Workbook.create();
    await Workbook.read(back, bytes);
    expect(Cell.getFormula(Workbook.getWorksheet(back, "S")!, "A2")).toBe(
      "SUM(Tb[[#This Row],[x]:[y]])"
    );
  });

  it("leaves the model text as the caller wrote it", async () => {
    const handle = Workbook.create();
    const ws = Workbook.addWorksheet(handle, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)" });
    await Workbook.toBuffer(handle);
    expect(Cell.getFormula(ws, "A1")).toBe("SEQUENCE(3)");
  });

  it("reads the stored form back as stored", async () => {
    // What every reader of an Excel-made workbook has always returned, and what the formula engine evaluates.
    const handle = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(handle, "S"), "A1", { formula: "XLOOKUP(2,B1:B4,B1:B4)" });
    const back = Workbook.create();
    await Workbook.read(back, await Workbook.toBuffer(handle));
    expect(Cell.getFormula(Workbook.getWorksheet(back, "S")!, "A1")).toBe(
      "_xlfn.XLOOKUP(2,B1:B4,B1:B4)"
    );
  });

  it("writes them from the streaming writer", async () => {
    const chunks: Uint8Array[] = [];
    const sink = new WritableStream<Uint8Array>({
      write(chunk) {
        chunks.push(chunk);
      }
    });
    const writer = new Stream.WorkbookWriter({ stream: sink as never });
    const ws = writer.addWorksheet("S");
    Stream.setCellValue(ws.getCell("A1"), { formula: "UNIQUE(B1:B4)" });
    ws.commit();
    await writer.commit();
    const bytes = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
    let at = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, at);
      at += chunk.length;
    }
    // `UNIQUE` produces an array, so it is a dynamic array — as the buffered writer stores it.
    expect((await parts(bytes)).sheet).toContain(
      '<c r="A1" cm="1"><f t="array" ref="A1">_xlfn.UNIQUE(B1:B4)</f>'
    );
  });
});
