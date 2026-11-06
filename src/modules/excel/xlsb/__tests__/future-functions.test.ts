/**
 * Functions the `Ftab` has no id for.
 *
 * `XLOOKUP`, `TEXTJOIN`, `CONFIDENCE.T`, `LET`, `SEQUENCE` — everything newer than the enumeration MS-XLSB
 * froze. These used to be refused by name, with the reason that they "have no iftab". True, and not the end of
 * it: Excel calls such a function through a **`PtgName`** naming a hidden `_xlfn.*` stub, with
 * `PtgFuncVar`'s `tab` set to `0x00FF` — "the callee is the first operand".
 *
 * Every constant below is read off Excel's own bytes rather than inferred:
 *
 * | Where                                   | Bytes                       | Meaning                          |
 * | --------------------------------------- | --------------------------- | -------------------------------- |
 * | `poi-bug66682.xlsb`, an array formula   | `23 02 00 00 00 42 01 ff 00`| `PtgName(2)`, `PtgFuncVar(1,255)`|
 * | Excel's repair of `scientific-analysis` | `flags 0x0002000b`, `1c 1d` | the stub's flags and `#NAME?` body|
 *
 * The specification then explains why those flags are not a choice: 2.4.674 requires `fFutureFunction` to be 0
 * unless `fHidden`, `fFunc` and `fProc` are all 1 and `fOB`, `fCalcExp`, `fgrp`, `fPublished`, `fBuiltin` and
 * `fWorkbookParam` are all 0, with a NULL comment and a global `itab`. `0x0002000b` is exactly that.
 */
import { Archive } from "@archive";
import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import { validateXlsbBuffer } from "@excel/utils/xlsb-validator";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import {
  FUTURE_FUNCTION_FLAGS,
  FUTURE_FUNCTION_STUB_RGCE,
  futureFunctionStubName,
  isFutureFunction
} from "@excel/xlsb/formula/ptg";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

/** A workbook whose A1 holds `formula`. */
function withFormula(formula: string): Workbook.Handle {
  const handle = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(handle, "S"), "A1", { formula, result: 1 } as never);
  return handle;
}

/** A workbook read from an XLSX whose A1 holds `formula` as an ordinary formula — no dynamic-array metadata. */
async function storedOrdinary(formula: string): Promise<Workbook.Handle> {
  const handle = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(handle, "S"), "A1", 1);
  const entries = new Map(
    [...(await extractAll(await Workbook.toBuffer(handle)))].map(([path, entry]) => [
      path,
      entry.data
    ])
  );
  const sheet = new TextDecoder()
    .decode(entries.get("xl/worksheets/sheet1.xml"))
    .replace(/<c r="A1"[^>]*>.*?<\/c>/, `<c r="A1"><f>${formula}</f><v>1</v></c>`);
  entries.set("xl/worksheets/sheet1.xml", new TextEncoder().encode(sheet));
  return Workbook.read(await Archive.packZip(entries));
}

/** The `BrtName` records and the A1 token stream of a workbook's XLSB form. */
async function parts(handle: Workbook.Handle): Promise<{
  readonly names: readonly { name: string; flags: number; rgce: Uint8Array }[];
  readonly rgce: Uint8Array;
  readonly problems: readonly string[];
}> {
  const bytes = await Workbook.toBuffer(handle, { format: "xlsb", unsupported: "ignore" });
  const archive = await extractAll(bytes);
  const names: { name: string; flags: number; rgce: Uint8Array }[] = [];
  for (const record of iterateInterpretableRecords(archive.get("xl/workbook.bin")!.data, "w")) {
    if (recordSpec(record.id)?.name !== "BrtName") {
      continue;
    }
    const view = new DataView(record.payload.buffer, record.payload.byteOffset);
    const characters = view.getUint32(9, true);
    let name = "";
    for (let index = 0; index < characters; index += 1) {
      name += String.fromCharCode(view.getUint16(13 + index * 2, true));
    }
    const at = 13 + characters * 2;
    const length = view.getUint32(at, true);
    names.push({
      name,
      flags: view.getUint32(0, true),
      rgce: record.payload.slice(at + 4, at + 4 + length)
    });
  }

  let rgce = new Uint8Array(0);
  for (const record of iterateInterpretableRecords(
    archive.get("xl/worksheets/sheet1.bin")!.data,
    "s"
  )) {
    if (!(recordSpec(record.id)?.name ?? "").startsWith("BrtFmla")) {
      continue;
    }
    const view = new DataView(record.payload.buffer, record.payload.byteOffset);
    rgce = record.payload.slice(22, 22 + view.getUint32(18, true));
  }

  const validation = await validateXlsbBuffer(bytes);
  return {
    names,
    rgce,
    problems: [...new Set((validation.problems ?? []).map(problem => problem.kind))]
  };
}

describe("which functions need a stub", () => {
  it.each(["XLOOKUP", "TEXTJOIN", "CONFIDENCE.T", "LET", "SEQUENCE", "IFS", "CONCAT"])(
    "treats %s as a future function",
    name => {
      expect(isFutureFunction(name)).toBe(true);
    }
  );

  it.each(["SUM", "IF", "AVERAGEIF", "COUNTIFS", "IFERROR", "MEDIAN"])(
    "treats %s as a built-in with an id",
    name => {
      // These have `Ftab` ids and must keep being called by id — writing one through a stub would produce a
      // formula that works and is not the one Excel writes.
      expect(isFutureFunction(name)).toBe(false);
    }
  );

  it("prefixes the stub the way Excel spells it", () => {
    expect(futureFunctionStubName("xlookup")).toBe("_xlfn.XLOOKUP");
  });
});

describe("a call to a future function", () => {
  it("writes the stub with the flags the specification requires", async () => {
    const found = await parts(withFormula("XLOOKUP(1,B1:B3,C1:C3)"));
    const stub = found.names.find(entry => entry.name === "_xlfn.XLOOKUP");
    expect(stub).toBeDefined();
    // `0x0002000b` — `fHidden | fFunc | fProc | fFutureFunction`, and Excel's own value.
    expect(stub!.flags).toBe(FUTURE_FUNCTION_FLAGS);
    expect(stub!.flags).toBe(0x0002000b);
    // `PtgErr(#NAME?)`, which is the body Excel gives every stub.
    expect([...stub!.rgce]).toEqual([...FUTURE_FUNCTION_STUB_RGCE]);
    expect([...stub!.rgce]).toEqual([0x1c, 0x1d]);
  });

  it("calls it through PtgName with iftab 0x00FF", async () => {
    // A call returning one value: one returning an array (`XLOOKUP`) is a dynamic array, whose cell holds a `PtgExp`.
    const found = await parts(withFormula("XMATCH(1,B1:B3,0)"));
    // First token: `PtgName` in the reference class, then the one-based name index.
    expect(found.rgce[0]).toBe(0x23);
    expect(found.rgce[1]).toBe(1);
    // Last four: `PtgFuncVar`, then `cparams` and `tab`. **`cparams` counts the name**, so three arguments
    // make four operands — the detail `poi-bug66682.xlsb` settles.
    const tail = found.rgce.slice(-4);
    expect(tail[0]).toBe(0x42);
    expect(tail[1]).toBe(4);
    expect(tail[2]! | (tail[3]! << 8)).toBe(0x00ff);
  });

  it("does not report a loss", async () => {
    await expect(
      Workbook.toBuffer(withFormula('TEXTJOIN(",",TRUE,A2:A4)'), {
        format: "xlsb",
        unsupported: "error"
      })
    ).resolves.toBeInstanceOf(Uint8Array);
  });

  it("passes the validator", async () => {
    expect((await parts(withFormula("SEQUENCE(3)"))).problems).toEqual([]);
  });

  it.each([
    "XLOOKUP(1,B1:B3,C1:C3)",
    "CONFIDENCE.T(0.05,2,3)",
    'TEXTJOIN(",",TRUE,B1:B3)',
    "SEQUENCE(3)"
  ])("reads %s back with the _xlfn. prefix the XLSX container needs", async formula => {
    // **The prefix is how the call is spelled outside the token stream.** Excel's own XLSX of a cell calling
    // `CONCAT` holds `_xlfn.CONCAT(…)`, and an unprefixed `CONCAT(…)` in XLSX is a user-defined function Excel
    // shows as `#NAME?`. Reading the stub's name without it made a workbook converted XLSB → XLSX lose every such
    // call; the name the stub carries is the name the formula has.
    const back = Workbook.create();
    await Workbook.read(
      back,
      await Workbook.toBuffer(withFormula(formula), { format: "xlsb", unsupported: "ignore" })
    );
    expect(
      (Cell.getValue(Workbook.listWorksheets(back)[0]!, "A1") as { formula?: string }).formula
    ).toBe(`_xlfn.${formula}`);
  });

  it.each(["_xlfn.CONCAT(B1,B2)", "_xlfn._xlws.SORT(B1:B3)", "_xlfn.STDEV.S(B1:B3)"])(
    "round-trips %s, already prefixed, unchanged",
    async formula => {
      // The form an XLSX read produces. It used to find no stub — the collector skipped prefixed names and the
      // encoder looked for `_xlfn._XLFN.CONCAT` — so the formula was refused and written as a blank.
      const handle = withFormula(formula);
      const { names } = await parts(handle);
      expect(names.map(entry => entry.name)).toContain(formula.slice(0, formula.indexOf("(")));
      const back = Workbook.create();
      await Workbook.read(back, await Workbook.toBuffer(handle, { format: "xlsb" }));
      expect(
        (Cell.getValue(Workbook.listWorksheets(back)[0]!, "A1") as { formula?: string }).formula
      ).toBe(formula);
    }
  );

  it("interns one stub for repeated calls", async () => {
    const handle = Workbook.create();
    const sheet = Workbook.addWorksheet(handle, "S");
    Cell.setValue(sheet, "A1", { formula: "SEQUENCE(3)", result: 1 } as never);
    Cell.setValue(sheet, "A2", { formula: "SEQUENCE(4)", result: 1 } as never);
    const found = await parts(handle);
    expect(found.names.filter(entry => entry.name === "_xlfn.SEQUENCE")).toHaveLength(1);
  });

  it("does not stub a built-in", async () => {
    const found = await parts(withFormula("SUM(A2:A4)"));
    expect(found.names.filter(entry => entry.name.startsWith("_xlfn."))).toEqual([]);
    // And the call is by id: `PtgFuncVar` with `SUM`'s `iftab`, not `0x00FF`.
    const tail = found.rgce.slice(-4);
    expect(tail[2]! | (tail[3]! << 8)).not.toBe(0x00ff);
  });
});

describe("the class of a future function's result", () => {
  /** The `PtgFuncVar` byte of the stub call — the one whose `tab` is `0x00FF`. */
  function stubCallClass(rgce: Uint8Array): number[] {
    const found: number[] = [];
    for (let at = 0; at + 3 < rgce.length; at++) {
      if ((rgce[at]! & 0x1f) === 0x02 && rgce[at + 2] === 0xff && rgce[at + 3] === 0x00) {
        found.push(rgce[at]! & 0x60);
      }
    }
    return found;
  }

  it.each([
    "SUM(FILTER(B1:B4,B1:B4>1))",
    "SUM(SEQUENCE(3))",
    "ROWS(SEQUENCE(5))",
    "SUM(VSTACK(1,2))",
    "SUM(UNIQUE(B1:B4))",
    "SUM(SORT(B1:B4))",
    "SUM(SORTBY(B1:B4,B1:B4))",
    "SUM(HSTACK(1,2))",
    "SUM(CHOOSECOLS(A1:B4,2))",
    "SUM(TOCOL(B1:B4))"
  ])("keeps an array-returning call whole as an argument: %s", async formula => {
    // In the value class Excel intersects the result to one cell — `SUM(FILTER(…))` computes 2 rather than 9.
    expect(stubCallClass((await parts(withFormula(formula))).rgce)).toEqual([0x60]);
  });

  // `TAKE` and `DROP` were asserted in the array class above, after LibreOffice's table; Excel's own XLSB of
  // `SUM(_xlfn.TAKE(…))` and `SUM(_xlfn.DROP(…))` has them in the reference class, as `TRIMRANGE` and `ANCHORARRAY`.
  it.each([
    "SUM(TAKE(B1:B4,2))",
    "SUM(DROP(B1:B4,1))",
    "SUM(TRIMRANGE(B1:B20))",
    "SUM(ANCHORARRAY(B1))"
  ])("keeps a reference-returning call a reference as an argument: %s", async formula => {
    // In the value class Excel shows `SUM(@TRIMRANGE(…))` and sums one cell.
    expect(stubCallClass((await parts(withFormula(formula))).rgce)).toEqual([0x20]);
  });

  it("keeps the value class for the whole formula", async () => {
    // An ordinary formula, as a file stores it: set through the API, `SEQUENCE(3)` is a dynamic array.
    expect(stubCallClass((await parts(await storedOrdinary("_xlfn.SEQUENCE(3)"))).rgce)).toEqual([
      0x40
    ]);
  });

  it("stubs a worksheet function with both prefixes, written bare or prefixed", async () => {
    for (const formula of ["SUM(FILTER(B1:B4,B1:B4>1))", "SUM(_xlfn.FILTER(B1:B4,B1:B4>1))"]) {
      const { names } = await parts(withFormula(formula));
      expect(names.map(entry => entry.name)).toEqual(["_xlfn._xlws.FILTER"]);
    }
  });
});

/**
 * `PY` and `PQSOURCE`: Excel refuses the package this writer produced for them (`lib-py.xlsb` in the fx21 checks), and
 * no Excel-written file shows a form it opens. They are reported, and with `unsupported: "ignore"` the cell keeps its
 * cached value and the package carries no stub for them.
 */
describe("functions with no XLSB form", () => {
  it.each(["PY(0,0)", "PQSOURCE(1)", "_xlfn._xlws.PY(0,0)"])("reports %s", async formula => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Cell.setValue(sheet, "A1", { formula, result: 7 } as never);
    Cell.setValue(sheet, "A2", { formula: "1+1" } as never);
    await expect(Workbook.toBuffer(workbook, { format: "xlsb" })).rejects.toThrow(
      /S!A1: formula \(a call to (PY|PQSOURCE) is not supported yet\)/
    );
    const bytes = await Workbook.toBuffer(workbook, { format: "xlsb", unsupported: "ignore" });
    const workbookPart = (await extractAll(bytes)).get("xl/workbook.bin")!.data;
    expect(new TextDecoder("utf-16le").decode(workbookPart)).not.toMatch(/PY|PQSOURCE/);
    const back = Workbook.create();
    await Workbook.read(back, bytes);
    const reread = Workbook.getWorksheet(back, "S")!;
    expect(Cell.getValue(reread, "A1")).toBe(7);
    expect(Cell.getFormula(reread, "A2")).toBe("1+1");
  });
});
