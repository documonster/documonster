/**
 * References through every token that carries one, checked against **MS-XLSB's field layout** rather than against
 * this codec's own encoder.
 *
 * The defect that prompted this file is the shape this module keeps producing: `readRelativeArea` read a
 * `RgceAreaRel` as two `RgceLocRel`s — `(row, col), (row, col)` — where the specification (2.5.98.90) lays it out
 * as `rowFirst, rowLast, columnFirst, columnLast`. The writer followed the specification, so `SUM($A$1:$B$2)` in a
 * conditional format came back `SUM($B$1:$B$1)` and a shared `SUM(A1:B1)` named row 4294836225 — and no test saw
 * it, because every relative case in the suite was a single cell.
 *
 * So the first half decodes bytes built by hand from the specification's tables, and the second half asks the
 * question a caller asks: does every place a formula can live give back, through XLSB, what XLSX gives back?
 */

import { extractAll } from "@archive/unzip/extract";
import { Cell, ConditionalFormatting, DefinedNames, Workbook, Worksheet } from "@excel";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { decodePtg, encodePtg, type PtgContext } from "@excel/xlsb/formula/ptg";
import { recordSpec } from "@excel/xlsb/spec/records";
import { parse } from "@formula/syntax/parser";
import { printAst } from "@formula/syntax/print";
import { tokenize } from "@formula/syntax/tokenizer";
import { describe, expect, it } from "vitest";

const CONTEXT: PtgContext = { sheetNames: ["S0", "S1"], definedNames: ["Rate"] };

/** Little-endian bytes for a token: `ptg`, then each field at its width. */
function token(ptg: number, ...fields: readonly (readonly [number, 2 | 4])[]): Uint8Array {
  const bytes: number[] = [ptg];
  for (const [value, width] of fields) {
    for (let index = 0; index < width; index++) {
      bytes.push((value >>> (8 * index)) & 0xff);
    }
  }
  return Uint8Array.from(bytes);
}

function decodeText(tokens: Uint8Array, origin?: { row: number; column: number }): string {
  const decoded = decodePtg(tokens, origin === undefined ? CONTEXT : { ...CONTEXT, origin }, "t");
  if ("sharedRow" in decoded) {
    throw new Error("decoded as PtgExp");
  }
  return printAst(decoded);
}

/** C5: row 4, column 2 — an origin with room for negative offsets on both axes. */
const C5 = { row: 4, column: 2 };
const ROW_REL = 0x8000;
const COL_REL = 0x4000;

describe("PtgAreaN is rowFirst, rowLast, columnFirst, columnLast (MS-XLSB 2.5.98.90)", () => {
  it("reads absolute corners as positions", () => {
    // `$A$1:$B$2`: rows 0 and 1, columns 0 and 1, no relative bit anywhere.
    expect(decodeText(token(0x2d, [0, 4], [1, 4], [0, 2], [1, 2]), C5)).toBe("$A$1:$B$2");
  });

  it("reads relative corners as offsets, in Excel's 20-bit spelling and the literal 32-bit one", () => {
    // `A1:B2` from C5: rows -4 and -3, columns -2 and -1.
    const both = ROW_REL | COL_REL;
    expect(
      decodeText(
        token(0x2d, [0xffffc, 4], [0xffffd, 4], [both | 0x3ffe, 2], [both | 0x3fff, 2]),
        C5
      )
    ).toBe("A1:B2");
    // `RwRelNeg` is a signed 32-bit integer (2.5.129), so -4 may also arrive as 0xFFFFFFFC.
    expect(
      decodeText(
        token(0x2d, [0xfffffffc, 4], [0xfffffffd, 4], [both | 0x3ffe, 2], [both | 0x3fff, 2]),
        C5
      )
    ).toBe("A1:B2");
  });

  it("decides each axis of each corner by its own flag", () => {
    // `$A1:B$2` from C5: first corner row-relative (-4), column absolute 0; last corner row absolute 1,
    // column-relative (-1).
    expect(
      decodeText(token(0x2d, [0xffffc, 4], [1, 4], [ROW_REL | 0, 2], [COL_REL | 0x3fff, 2]), C5)
    ).toBe("$A1:B$2");
  });

  it("recognises a whole column whose rows are pinned and whose column is relative", () => {
    expect(
      decodeText(token(0x2d, [0, 4], [0xfffff, 4], [COL_REL | 0, 2], [COL_REL | 0, 2]), C5)
    ).toBe("C:C");
  });
});

describe("RgceLocRel resolution (MS-XLSB 2.5.98.92)", () => {
  it("wraps an offset above row 1 to the bottom of the sheet, and left of A to XFD", () => {
    const both = ROW_REL | COL_REL;
    expect(decodeText(token(0x4c, [0xfffff, 4], [both, 2]), { row: 0, column: 0 })).toBe(
      "A1048576"
    );
    expect(decodeText(token(0x4c, [0, 4], [both | 0x3fff, 2]), { row: 0, column: 0 })).toBe("XFD1");
  });

  it("does not sign-extend an absolute row", () => {
    // `$A$600000` in a template: row 599999 is above 0x80000, so treating it as a 20-bit offset made it negative.
    expect(decodeText(token(0x4c, [599999, 4], [0, 2]), C5)).toBe("$A$600000");
  });
});

describe("whole rows and columns keep their $ markers", () => {
  for (const source of ["SUM(A:A)", "SUM($A:$A)", "SUM($A:C)", "SUM(1:1)", "SUM($1:$3)"]) {
    it(`${source} in a cell`, () => {
      const tokens = encodePtg(parse(tokenize(source)), CONTEXT, "A1");
      expect(decodeText(tokens)).toBe(source);
    });
    it(`${source} in a template`, () => {
      const tokens = encodePtg(
        parse(tokenize(source)),
        { ...CONTEXT, origin: C5, relativeToOrigin: true },
        "C5"
      );
      expect(decodeText(tokens, C5)).toBe(source);
    });
  }

  it("reads Excel's spelling of A:A — rows absolute, column relative", () => {
    // Apache POI, which is written to match Excel, builds `A:B` as rows 0..max row-absolute with the columns'
    // `$` taken from the text.
    expect(decodeText(token(0x25, [0, 4], [0xfffff, 4], [COL_REL | 0, 2], [COL_REL | 1, 2]))).toBe(
      "A:B"
    );
  });
});

describe("PtgArea3d and the error forms, from an Excel-written file", () => {
  // `cal-issue_182.xlsb` (calamine's corpus, written by Excel), `xl/workbook.bin`, two `BrtName` definitions.
  it("reads PtgArea3d as rowFirst, rowLast, columnFirst, columnLast", () => {
    // `MyDataTypes`: `3b 00 00 00 00 00 00 05 00 00 00 00 00 00 00` — ixti 0, rows 0..5, column 0 twice.
    // A sheet named like a cell (`S0`, `S1`) is quoted, as Excel writes it: `S1!A1` is not a sheet
    // reference. The shared rule in @utils/sheet-name; these used to print unquoted.
    expect(
      decodeText(Uint8Array.of(0x3b, 0x00, 0x00, 0, 0, 0, 0, 5, 0, 0, 0, 0x00, 0x00, 0x00, 0x00))
    ).toBe("'S0'!$A$1:$A$6");
  });

  it("reads PtgAreaErr3d and PtgRefErr3d as #REF! on their sheet instead of refusing the formula", () => {
    // `MyBrokenRange`: `3d 01 00 06 00 00 00 06 00 00 00 00 00 03 00` — `0x3d` is PtgAreaErr3d (0x1D, reference
    // class), whose twelve bytes after the ixti are unused (MS-XLSB 2.5.98.21).
    //
    // Qualified with the sheet the `ixti` names, because that is the text Excel shows and stores in XLSX for the
    // same token: `Other!#REF!` is `5c 03 00 …` in Excel's own output.
    expect(
      decodeText(Uint8Array.of(0x3d, 0x01, 0x00, 6, 0, 0, 0, 6, 0, 0, 0, 0x00, 0x00, 0x03, 0x00))
    ).toBe("'S1'!#REF!");
    // PtgRefErr3d: ixti, then six unused bytes (2.5.98.71).
    expect(decodeText(Uint8Array.of(0x3c, 0, 0, 0, 0, 0, 0, 0, 0))).toBe("'S0'!#REF!");
  });
});

describe("a template refuses what its record cannot hold", () => {
  it("rejects a sheet-qualified reference, which SharedParsedFormula and CFParsedFormula forbid", () => {
    for (const source of ["S1!A1+1", "SUM(S1!A1:B2)", "SUM(S1!A:A)"]) {
      expect(
        () =>
          encodePtg(
            parse(tokenize(source)),
            { ...CONTEXT, origin: C5, relativeToOrigin: true },
            "C5"
          ),
        source
      ).toThrow(/another sheet/);
    }
  });
});

// =============================================================================
// Through the package: XLSB must report what XLSX reports
// =============================================================================

/** Formulas that are legal anywhere a single-sheet template is. */
const LOCAL = [
  "A1",
  "$A$1",
  "$A1",
  "A$1",
  "A1+B2",
  "SUM(A1:B2)",
  "SUM($A$1:$B$2)",
  "SUM($A1:B$2)",
  "SUM(A$1:$A1)",
  "COUNTIF($C$5:$D$9,C5)",
  "SUM(A:A)",
  "SUM($A:$A)",
  "SUM(A:C)",
  "SUM(1:1)",
  "SUM($1:$2)",
  "Rate*2",
  "IFERROR(A1,#N/A)",
  "#REF!+1",
  "A1048576",
  "XFD1",
  "$A$600000"
];

/** Formulas a cell, an array formula or a name may hold but a template may not. */
const QUALIFIED = ["Other!A1", "SUM(Other!$A$1:B2)", "SUM(Other!A:A)", "SUM(Other!$1:$1)"];

function workbookWith(build: (sheet: ReturnType<typeof Workbook.addWorksheet>) => void) {
  const handle = Workbook.create();
  const sheet = Workbook.addWorksheet(handle, "S");
  Workbook.addWorksheet(handle, "Other");
  DefinedNames.add(Workbook.getDefinedNames(handle), "S!$Z$1", "Rate");
  build(sheet);
  return handle;
}

async function reopen(
  handle: ReturnType<typeof Workbook.create>,
  format: "xlsx" | "xlsb"
): Promise<ReturnType<typeof Workbook.addWorksheet>> {
  const bytes = await Workbook.toBuffer(handle, { format });
  const copy = Workbook.create();
  await Workbook.read(copy, bytes, { format });
  return Workbook.listWorksheets(copy)[0]!;
}

/** Hex of every record called `name` in the first sheet part, written as XLSB. */
async function recordsNamed(
  handle: ReturnType<typeof Workbook.create>,
  name: string
): Promise<string[]> {
  const parts = await extractAll(await Workbook.toBuffer(handle, { format: "xlsb" }));
  const hex: string[] = [];
  for (const entry of iterateInterpretableRecords(
    parts.get("xl/worksheets/sheet1.bin")!.data,
    "s"
  )) {
    if (recordSpec(entry.id)?.name === name) {
      hex.push([...entry.payload].map(byte => byte.toString(16).padStart(2, "0")).join(" "));
    }
  }
  return hex;
}

async function bothWays<T>(
  build: (sheet: ReturnType<typeof Workbook.addWorksheet>) => void,
  look: (sheet: ReturnType<typeof Workbook.addWorksheet>) => T
): Promise<{ xlsx: T; xlsb: T }> {
  return {
    xlsx: look(await reopen(workbookWith(build), "xlsx")),
    xlsb: look(await reopen(workbookWith(build), "xlsb"))
  };
}

describe("every formula container agrees with XLSX", () => {
  it.each([...LOCAL, ...QUALIFIED])("a cell formula %s", async formula => {
    const { xlsx, xlsb } = await bothWays(
      sheet => Cell.setValue(sheet, "C5", { formula, result: 0 }),
      sheet => Cell.getFormula(sheet, "C5")
    );
    expect(xlsb).toBe(xlsx);
    expect(xlsb).toBe(formula);
  });

  it.each(LOCAL)("a shared formula %s, master and every follower", async formula => {
    const cells = ["C5", "D5", "C6", "D9"];
    const { xlsx, xlsb } = await bothWays(
      sheet => Worksheet.fillFormula(sheet, "C5:D9", formula),
      sheet => cells.map(address => Cell.getFormula(sheet, address))
    );
    expect(xlsb).toEqual(xlsx);
    expect(xlsb[0]).toBe(formula);
  });

  it.each(QUALIFIED)(
    "a shared formula %s that names another sheet, which is written cell by cell",
    async formula => {
      const cells = ["C5", "D5", "C6", "D9"];
      const { xlsx, xlsb } = await bothWays(
        sheet => Worksheet.fillFormula(sheet, "C5:D9", formula),
        sheet => cells.map(address => Cell.getFormula(sheet, address))
      );
      expect(xlsb).toEqual(xlsx);
      // No `BrtShrFmla` at all: `SharedParsedFormula` may not hold a `PtgRef3d` or `PtgArea3d`.
      const handle = workbookWith(sheet => Worksheet.fillFormula(sheet, "C5:D9", formula));
      expect(await recordsNamed(handle, "BrtShrFmla")).toEqual([]);
    }
  );

  it.each([...LOCAL, ...QUALIFIED])("an array formula %s", async formula => {
    const { xlsx, xlsb } = await bothWays(
      sheet =>
        Cell.setValue(sheet, "C5", {
          formula,
          shareType: "array",
          ref: "C5:D6",
          result: 0
        } as never),
      sheet => Cell.getFormula(sheet, "C5")
    );
    expect(xlsb).toBe(xlsx);
  });

  it.each(LOCAL)("a conditional format %s", async formula => {
    const { xlsx, xlsb } = await bothWays(
      sheet =>
        ConditionalFormatting.add(sheet, {
          ref: "C5:D9",
          rules: [
            { type: "expression", formulae: [`${formula}>0`], style: { font: { bold: true } } }
          ]
        }),
      sheet => Worksheet.getModel(sheet).conditionalFormattings?.[0]?.rules[0]
    );
    expect((xlsb as { formulae?: unknown } | undefined)?.formulae).toEqual(
      (xlsx as { formulae?: unknown } | undefined)?.formulae
    );
    expect((xlsb as { formulae?: unknown } | undefined)?.formulae).toEqual([`${formula}>0`]);
  });

  it.each(LOCAL)("a custom validation %s", async formula => {
    const { xlsx, xlsb } = await bothWays(
      sheet => Cell.setValidation(sheet, "C5", { type: "custom", formulae: [`${formula}>0`] }),
      sheet => (Cell.getValidation(sheet, "C5") as { formulae?: unknown } | undefined)?.formulae
    );
    expect(xlsb).toEqual(xlsx);
  });

  it.each([
    "S!$A$1",
    "S!$A$1:$B$2",
    "S!$A:$A",
    "S!$1:$2",
    "Other!$C$3:$D$4",
    "S!A1",
    "S!A1:B2",
    "S!A:A"
  ])("a defined name %s", async reference => {
    const target = async (format: "xlsx" | "xlsb") => {
      const handle = workbookWith(() => {});
      DefinedNames.add(Workbook.getDefinedNames(handle), reference, "Target");
      const copy = Workbook.create();
      await Workbook.read(copy, await Workbook.toBuffer(handle, { format }), { format });
      return DefinedNames.list(Workbook.getDefinedNames(copy)).find(
        entry => entry.name === "Target"
      );
    };
    expect(await target("xlsb")).toEqual(await target("xlsx"));
  });
});

describe("the bytes a template is written with", () => {
  it("lays a PtgAreaN out as rows, then columns", async () => {
    const handle = workbookWith(sheet =>
      ConditionalFormatting.add(sheet, {
        ref: "C5:D9",
        rules: [
          { type: "expression", formulae: ["SUM(A$1:B$2)>0"], style: { font: { bold: true } } }
        ]
      })
    );
    const hex = await recordsNamed(handle, "BrtBeginCFRule");
    // `PtgAreaN` (reference class): rowFirst 0, rowLast 1 — absolute — then the two relative columns as offsets
    // from C: -2 and -1, each with `fColRel` set.
    expect(hex[0]).toContain("2d 00 00 00 00 01 00 00 00 fe 7f ff 7f");
    // A fully absolute area in the same template is an ordinary `PtgArea`, as Excel writes `MAX($A$1:$A$10)`.
    const absolute = await recordsNamed(
      workbookWith(sheet =>
        ConditionalFormatting.add(sheet, {
          ref: "C5:D9",
          rules: [
            { type: "expression", formulae: ["SUM($A$1:$B$2)>0"], style: { font: { bold: true } } }
          ]
        })
      ),
      "BrtBeginCFRule"
    );
    expect(absolute[0]).toContain("25 00 00 00 00 01 00 00 00 00 00 01 00");
  });
});
