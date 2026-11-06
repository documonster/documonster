/**
 * An immediately invoked `LAMBDA` — `LAMBDA(n,n*2)(A3)` — through the calculation engine and both containers.
 *
 * XLSX stores formula text, and Excel's spelling of the call is `_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2)(A3)`: the applied
 * argument list follows the call verbatim. XLSB stores a token stream; MS-XLSB documents no token for calling the
 * result of an expression, and Excel's own file shows what it writes — the `LAMBDA` call, the applied arguments, then
 * `PtgFuncVar` with `iftab` `0x00FF` and the callee counted as its first operand. Byte-for-byte agreement with Excel's
 * file is pinned in `future-function-tokens.node.test.ts`; here the round trip and the calculation.
 */
import { extractAll } from "@archive/unzip/extract";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("an applied argument list", () => {
  it("evaluates an immediately invoked, a curried and a LET-bound lambda", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A3", 4);
    Cell.setValue(ws, "B1", { formula: "LAMBDA(n,n*2)(A3)" });
    Cell.setValue(ws, "B2", { formula: "LAMBDA(x,LAMBDA(y,x+y))(1)(2)" });
    Cell.setValue(ws, "B3", { formula: "LET(f,LAMBDA(x,LAMBDA(y,x*y)),f(3)(5))" });
    Cell.setValue(ws, "B4", { formula: "1+LAMBDA(n,n)(2)*3" });
    Cell.setValue(ws, "B5", { formula: "LAMBDA(7)()" });
    Cell.setValue(ws, "B6", { formula: "SUM(A3:A3,1)(2)" });
    Cell.setValue(ws, "B7", { formula: "LAMBDA(n,n)(1,2)" });
    Cell.setValue(ws, "B8", { formula: "LAMBDA(n,LET(m,n+1,LAMBDA(k,m*k)))(1)(10)" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(8);
    expect(Cell.getResult(ws, "B2")).toBe(3);
    expect(Cell.getResult(ws, "B3")).toBe(15);
    expect(Cell.getResult(ws, "B4")).toBe(7);
    expect(Cell.getResult(ws, "B5")).toBe(7);
    // Not a lambda, and the wrong number of arguments: both #VALUE!, as for a LET-bound call.
    expect(Cell.getResult(ws, "B6")).toEqual({ error: "#VALUE!" });
    expect(Cell.getResult(ws, "B7")).toEqual({ error: "#VALUE!" });
    expect(Cell.getResult(ws, "B8")).toBe(20);
  });

  it("is written to XLSX as Excel spells it, and reads back", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A3", 4);
    Cell.setValue(ws, "B1", { formula: "LAMBDA(n,n*2)(A3)" });
    Cell.setValue(ws, "B2", { formula: "LAMBDA(x,LAMBDA(y,x+y))(1)(2)" });
    const bytes = await Workbook.toBuffer(wb);
    const sheet = new TextDecoder().decode(
      (await extractAll(bytes)).get("xl/worksheets/sheet1.xml")!.data
    );
    expect(sheet).toContain("<f>_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2)(A3)</f>");
    expect(sheet).toContain(
      "<f>_xlfn.LAMBDA(_xlpm.x,_xlfn.LAMBDA(_xlpm.y,_xlpm.x+_xlpm.y))(1)(2)</f>"
    );

    const back = Workbook.create();
    await Workbook.read(back, bytes);
    const read = Workbook.getWorksheet(back, "S")!;
    calculateFormulas(back);
    expect(Cell.getResult(read, "B1")).toBe(8);
    expect(Cell.getResult(read, "B2")).toBe(3);
  });

  it("is written to XLSB as Excel writes it, and reads back", async () => {
    // This used to be refused, while no Excel-written sample of the encoding existed; this test pinned that.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A3", 4);
    Cell.setValue(ws, "B1", { formula: "LAMBDA(n,n*2)(A3)" });
    Cell.setValue(ws, "B2", { formula: "LAMBDA(x,LAMBDA(y,x+y))(1)(2)" });
    Cell.setValue(ws, "B3", { formula: "SUM(LAMBDA(n,SEQUENCE(n))(3))" });
    const bytes = await Workbook.toBuffer(wb, { format: "xlsb" });
    const back = Workbook.create();
    await Workbook.read(back, bytes);
    const read = Workbook.getWorksheet(back, "S")!;
    expect(Cell.getFormula(read, "B1")).toBe("_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2)(A3)");
    expect(Cell.getFormula(read, "B2")).toBe(
      "_xlfn.LAMBDA(_xlpm.x,_xlfn.LAMBDA(_xlpm.y,_xlpm.x+_xlpm.y))(1)(2)"
    );
    calculateFormulas(back);
    expect(Cell.getResult(read, "B1")).toBe(8);
    expect(Cell.getResult(read, "B2")).toBe(3);
    expect(Cell.getResult(read, "B3")).toBe(6);
  });

  it("calls a LET-bound lambda in the reference class, so an array result is not intersected", async () => {
    // Excel's own `_xlfn.LET(_xlpm.f,…,_xlpm.f(1,2))` ends `22 03 ff 00 42 04 ff 00`. Written in the value class,
    // Excel shows `SUM(@f(3))` and computes 1 where the answer is 6.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "LET(f,LAMBDA(x,SEQUENCE(x)),SUM(f(3)))" });
    const archive = await extractAll(await Workbook.toBuffer(wb, { format: "xlsb" }));
    const sheet = archive.get("xl/worksheets/sheet1.bin")!.data;
    const hex = [...sheet].map(byte => byte.toString(16).padStart(2, "0")).join("");
    // `f(3)` is `23 <f> 1e 03 00 22 02 ff 00`.
    expect(hex).toContain("1e03002202ff00");
    expect(hex).not.toContain("1e03004202ff00");
  });
});
