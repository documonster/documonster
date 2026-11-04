import { extractAll } from "@archive/unzip/extract";
/**
 * Bytes Excel wrote, transcribed: the cases each defect in this codec was found by, readable without the fixtures.
 *
 * Every hex string below was read off an XLSB that Excel (Microsoft Excel for Mac, 16.x) saved from XLSX text. The
 * encoder must produce exactly it, and the decoder must read it back to the text. `operand-classes.node.test.ts`
 * does the same over every cell of the Excel-made fixtures; this file names the cases.
 */
import { ConditionalFormatting, Cell, Workbook } from "@excel";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { encodeFormula, formulaText, type PtgContext } from "@excel/xlsb/formula/ptg";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const hex = (bytes: Uint8Array): string =>
  [...bytes].map(byte => byte.toString(16).padStart(2, "0")).join(" ");
const bytes = (text: string): Uint8Array =>
  Uint8Array.from(text.split(" ").map(pair => Number.parseInt(pair, 16)));

const CONTEXT: PtgContext = { sheetNames: ["S"], externSheets: [{ first: 0, last: 0 }] };

/** `[formula, Excel's rgce]`. */
const CASES: readonly (readonly [string, string])[] = [
  // Operator operands are values: `24` (reference class) here is `#VALUE!` in Excel.
  ["A1+B1", "44 00 00 00 00 00 c0 44 00 00 00 00 01 c0 03"],
  ["-A1", "44 00 00 00 00 00 c0 13"],
  ["A1&B1", "44 00 00 00 00 00 c0 44 00 00 00 00 01 c0 08"],
  ["ABS(A1-B1)", "44 00 00 00 00 00 c0 44 00 00 00 00 01 c0 04 41 18 00"],
  ["A1", "44 00 00 00 00 00 c0"],
  ["A1:A3", "45 00 00 00 00 02 00 00 00 00 c0 00 c0"],
  // References where a function takes them; arrays where it takes arrays.
  ["SUM(A1:A3,B1)", "25 00 00 00 00 02 00 00 00 00 c0 00 c0 24 00 00 00 00 01 c0 42 02 04 00"],
  [
    "SUMPRODUCT(A1:A3*B1:B3)",
    "65 00 00 00 00 02 00 00 00 00 c0 00 c0 65 00 00 00 00 02 00 00 00 01 c0 01 c0 05 42 01 e4 00"
  ],
  [
    "SUM(OFFSET(A1,1,1,2,2))",
    "19 01 00 00 24 00 00 00 00 00 c0 1e 01 00 1e 01 00 1e 02 00 1e 02 00 22 05 4e 00 19 10 00 00"
  ],
  [
    "A1:INDEX(A1:C3,2,2)",
    "49 1f 00 24 00 00 00 00 00 c0 25 00 00 00 00 02 00 00 00 00 c0 02 c0 1e 02 00 1e 02 00 22 03 1d 00 11"
  ],
  // `PtgAttrSum`, `PtgAttrSemi`, and the jump tables.
  ["SUM(A1:A3)", "25 00 00 00 00 02 00 00 00 00 c0 00 c0 19 10 00 00"],
  ["NOW()", "19 01 00 00 41 4a 00"],
  [
    'IF(A1>1,"a","b")',
    "44 00 00 00 00 00 c0 1e 01 00 0d 19 02 09 00 17 01 00 61 00 19 08 0c 00 17 01 00 62 00 19 08 03 00 42 03 01 00"
  ],
  [
    "CHOOSE(2,A1,B1,C1)",
    "1e 02 00 19 04 03 00 08 00 13 00 1e 00 29 00 24 00 00 00 00 00 c0 19 08 19 00 24 00 00 00 00 01 c0 19 08 0e 00 24 00 00 00 00 02 c0 19 08 03 00 42 04 64 00"
  ],
  ["IFERROR(A1/0,0)", "44 00 00 00 00 00 c0 1e 00 00 06 19 80 03 00 1e 00 00 19 08 02 00 41 e0 01"],
  // Display parentheses and the author's spacing.
  ["A1*(B1+C1)", "44 00 00 00 00 00 c0 44 00 00 00 00 01 c0 44 00 00 00 00 02 c0 03 15 05"],
  [
    "( A1+B1 )*2",
    "19 40 00 01 44 00 00 00 00 00 c0 44 00 00 00 00 01 c0 03 19 40 04 01 15 1e 02 00 05"
  ],
  [
    "SUM( A1, B1 )",
    "19 40 00 01 24 00 00 00 00 00 c0 19 40 00 01 24 00 00 00 00 01 c0 19 40 04 01 42 02 04 00"
  ],
  ["A1\n+B1", "44 00 00 00 00 00 c0 44 00 00 00 00 01 c0 19 40 01 01 03"],
  // A negated literal is one constant.
  ["-5", "1f 00 00 00 00 00 00 14 c0"],
  // `#REF!` typed is a reference error.
  ["#REF!", "4a 00 00 00 00 00 00"],
  // `EXACT` and `LENB` are fixed-arity.
  ["EXACT(D1,D1)", "44 00 00 00 00 03 c0 44 00 00 00 00 03 c0 41 75 00"],
  ['LENB("ab")', "17 02 00 61 00 62 00 41 d3 00"]
];

describe("the tokens Excel writes", () => {
  it.each(CASES)("%j", (formula, excel) => {
    expect(hex(encodeFormula(formula, CONTEXT, "test").rgce)).toBe(excel);
    expect(formulaText(bytes(excel), CONTEXT, "test")).toBe(formula);
  });

  it("writes _xlfn.SINGLE — XLSX's spelling of @ — as Excel does: the operand bare where that intersects it", () => {
    // Excel's XLSB of `_xlfn.SINGLE(A1:A3)` is the bare value-class area, which it shows as `@A1:A3`; inside `SUM`,
    // where the bare area would be summed whole, the call through the stub (`single-tokens.node.test.ts`).
    const context = { ...CONTEXT, definedNames: ["_xlfn.SINGLE"] };
    expect(hex(encodeFormula("_xlfn.SINGLE(A1:A3)", context, "test").rgce)).toBe(
      "45 00 00 00 00 02 00 00 00 00 c0 00 c0"
    );
    expect(hex(encodeFormula("SUM(_xlfn.SINGLE(A1:A3))", context, "test").rgce)).toBe(
      "23 01 00 00 00 25 00 00 00 00 02 00 00 00 00 c0 00 c0 22 02 ff 00 19 10 00 00"
    );
  });

  it("writes and reads an array constant's values in RgbExtra", () => {
    const encoded = encodeFormula('{1,"a";TRUE,#N/A}', CONTEXT, "test");
    // The token's padding is unused; Excel's own carries the column and row counts less one in front.
    expect(hex(encoded.rgce)).toBe("40 01 00 01 00 00 00 00 00 00 00 00 00 00 00");
    expect(hex(encoded.extra)).toBe(
      "02 00 00 00 02 00 00 00 00 00 00 00 00 00 00 f0 3f 01 01 00 61 00 02 01 04 2a 00 00 00"
    );
    expect(formulaText(encoded.rgce, CONTEXT, "test", encoded.extra)).toBe('{1,"a";TRUE,#N/A}');
  });

  it("reads a call through a future-function stub with its _xlfn. prefix", () => {
    // `_xlfn.CONCAT(A1,B1)` as Excel wrote it: `PtgName` 2, the arguments, `PtgFuncVar` 3 through 0x00FF.
    const context = { ...CONTEXT, definedNames: ["_xlfn._xlws.SORT", "_xlfn.CONCAT"] };
    const excel = "23 02 00 00 00 24 00 00 00 00 00 c0 24 00 00 00 00 01 c0 42 03 ff 00";
    expect(formulaText(bytes(excel), context, "test")).toBe("_xlfn.CONCAT(A1,B1)");
    expect(hex(encodeFormula("_xlfn.CONCAT(A1,B1)", context, "test").rgce)).toBe(excel);
  });

  it("reads and writes another sheet's local name through PtgNameX", () => {
    // `Other!LocalO` on sheet S: `ixti` 2 is the workbook-level entry (-2, -2), name 11.
    const names = Array.from({ length: 11 }, (_unused, index) => `N${index}`);
    names[10] = "LocalO";
    const context: PtgContext = {
      sheetNames: ["S", "Other"],
      externSheets: [
        { first: 0, last: 0 },
        { first: 1, last: 1 },
        { first: -2, last: -2 }
      ],
      definedNames: names,
      definedNameScopes: names.map((_name, index) => (index === 10 ? 1 : undefined)),
      sheetIndex: 0
    };
    expect(formulaText(bytes("59 02 00 0b 00 00 00"), context, "test")).toBe("Other!LocalO");
    expect(hex(encodeFormula("Other!LocalO", context, "test").rgce)).toBe("59 02 00 0b 00 00 00");
  });
});

/** The `Rgce` of every `BrtBeginCFRule` / `BrtDVal` in a workbook's first sheet, as hex. */
async function templateTokens(handle: Workbook.Handle): Promise<string[]> {
  const archive = await extractAll(await Workbook.toBuffer(handle, { format: "xlsb" }));
  const out: string[] = [];
  for (const record of iterateInterpretableRecords(
    archive.get("xl/worksheets/sheet1.bin")!.data,
    "sheet"
  )) {
    const name = recordSpec(record.id)?.name;
    if (name === "BrtBeginCFRule" || name === "BrtDVal") {
      out.push(hex(record.payload));
    }
  }
  return out;
}

describe("a multi-range rule is anchored at its bounding box, as Excel anchors it", () => {
  it("in a conditional format", async () => {
    const handle = Workbook.create();
    const sheet = Workbook.addWorksheet(handle, "S");
    Cell.setValue(sheet, "A1", 1);
    ConditionalFormatting.add(sheet, {
      ref: "H6:H8 G2:G4",
      rules: [{ type: "expression", formulae: ["H6>A6"], style: { font: { bold: true } } }]
    });
    // Excel: `4c 04 00 00 00 01 c0 4c 04 00 00 00 fa ff 0d` — H6 and A6 as offsets from G2. From H6, the first
    // range, A6 is seven columns left and wrapped to XFD: what Excel showed as `G2>XFD2`.
    expect((await templateTokens(handle))[0]).toContain(
      "4c 04 00 00 00 01 c0 4c 04 00 00 00 fa ff 0d"
    );
  });

  it("in a data validation", async () => {
    const handle = Workbook.create();
    const sheet = Workbook.addWorksheet(handle, "S");
    Cell.setValue(sheet, "A1", 1);
    // Each entry's formula is relative to its own cell, so the one rule Excel states as `A:A` from B4 is `C:C` at D4
    // and `A:A` at B12 and B13. This case used to give all three `A:A`, which only produced one record because the
    // writer merged rules by text — the same merge that made a relative reference mean different cells after a
    // round trip.
    Cell.setValidation(sheet, "D4", { type: "list", formulae: ["C:C"] });
    for (const address of ["B12", "B13"]) {
      Cell.setValidation(sheet, address, { type: "list", formulae: ["A:A"] });
    }
    // Excel: `2d 00 00 00 00 ff ff 0f 00 ff 7f ff 7f` — `A:A` as a column offset of -1 from B, the bounding box's
    // left edge. Written as positions, Excel read the list as `D:D`.
    expect((await templateTokens(handle))[0]).toContain("2d 00 00 00 00 ff ff 0f 00 ff 7f ff 7f");
  });
});
