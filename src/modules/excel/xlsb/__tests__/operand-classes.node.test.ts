/**
 * The tokens this encoder writes, against the tokens Excel wrote for the same text.
 *
 * Every formula cell in the two Excel-made fixtures (`__tests__/data/*.excel.xlsb`) is re-encoded from the text the
 * XLSX twin holds and compared byte for byte with what Excel stored. This is the check a round trip cannot make: an
 * encoder and a decoder that share a mistake agree with each other, and only Excel's bytes disagree. It is how the
 * operand classes were found wrong — `=A1+B1` written as `24 … 24 … 03`, reference class, which Excel evaluates as
 * `#VALUE!` — and it pins everything the encoder chooses: classes, `PtgAttrSum`, the `IF`/`CHOOSE`/`IFERROR` jump
 * offsets, `PtgAttrSemi`, `PtgParen`, the spacing tokens, `PtgArray` and its values, `PtgNameX`, the 3D forms.
 *
 * The decode direction is asserted for the same cells: Excel's bytes read back to Excel's text.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import { encodeCol } from "@excel/utils/address";
import { XlsbReader, iterateInterpretableRecords, readCell } from "@excel/xlsb/binary";
import { encodeFormula, formulaText } from "@excel/xlsb/formula/ptg";
import { readWorkbookPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";

/**
 * Where this encoder deliberately writes something else, by formula text — each a choice the specification
 * leaves open, and each still read back by Excel as the same formula computing the same value (the acceptance
 * matrix opens both forms in Excel and compares).
 *
 * - **A reference operator's mem token.** Excel caches the result of `A1:A3 A2:C2` or `(A1,B1)` in a `PtgMemArea`
 *   (or `PtgMemErr` when it is empty) with the areas in `RgbExtra`; this writer emits `PtgMemFunc`, which asks for
 *   evaluation rather than asserting a cached answer, for every such expression.
 */
const MEM_TOKEN_CHOICE = new Set([
  "A1:A3  B2:C2",
  "SUM(A1:A3 A2:C2)",
  "SUM((A1,B1))",
  "INDEX((A1:A3,B1:B3),1,1,2)",
  "AREAS((A1,B1))",
  "SUM(A1:B2 B2:C3)",
  "SUM((A1:A2,C1:C2))",
  "INDEX((A1:A3,B1:B3),2,1,2)",
  "AREAS((A1,B1,C1))",
  "A1:A3 A2:C2"
]);

interface ExcelCell {
  readonly sheet: number;
  readonly address: string;
  readonly row: number;
  readonly column: number;
  readonly rgce: Uint8Array;
  readonly extra: Uint8Array;
}

/** Every formula cell of an XLSB sheet part, with its `Rgce` and `RgbExtra`. */
function formulaCells(bytes: Uint8Array, sheet: number): ExcelCell[] {
  const cells: ExcelCell[] = [];
  let row = 0;
  for (const record of iterateInterpretableRecords(bytes, "sheet")) {
    const name = recordSpec(record.id)?.name;
    if (name === "BrtRowHdr") {
      row = new DataView(record.payload.buffer, record.payload.byteOffset).getUint32(0, true);
      continue;
    }
    if (
      name !== "BrtFmlaNum" &&
      name !== "BrtFmlaString" &&
      name !== "BrtFmlaBool" &&
      name !== "BrtFmlaError"
    ) {
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
    if (rgce[0] === 0x01) {
      continue; // `PtgExp`: a member of an array formula, whose body is its master's.
    }
    cells.push({ sheet, address: `${encodeCol(column)}${row + 1}`, row, column, rgce, extra });
  }
  return cells;
}

/**
 * The bytes the specification calls unused, which Excel fills with whatever was in memory: the last ten of
 * `PtgArray`'s fourteen, and the two after a leading `PtgAttrSemi` (`19 01 05 00` in front of `_xlfn.SHEET`).
 */
function maskUnused(rgce: Uint8Array, hasArray: boolean): string {
  const hex = [...rgce].map(byte => byte.toString(16).padStart(2, "0"));
  if (rgce[0] === 0x19 && rgce[1] === 0x01) {
    hex[2] = "..";
    hex[3] = "..";
  }
  // Only where the text has an array constant, so a `0x40` inside some other token's payload masks nothing.
  for (let at = 0; hasArray && at < hex.length; at++) {
    if ((rgce[at]! === 0x20 || rgce[at]! === 0x40 || rgce[at]! === 0x60) && at + 15 <= hex.length) {
      for (let pad = at + 5; pad < at + 15; pad++) {
        hex[pad] = "..";
      }
    }
  }
  return hex.join(" ");
}

// `formula-classes`: an expression handed to each parameter LibreOffice's table converts as "original" and that accepts
// one (`ROWS(A1:A5+0)`, `VLOOKUP(1,A1:A5+0,1)`, `CHOOSE(A1:A5+0,1,2)`, …), opened by Excel 16 and saved in both formats.
describe.each(["formula-matrix", "formula-edges", "formula-classes"])(
  "%s: tokens Excel wrote",
  fixture => {
    it("are what this encoder writes for the same text, and read back to that text", async () => {
      const archive = await extractAll(await readFile(`${DIRECTORY}/${fixture}.excel.xlsb`));
      const { formulaContext, sheetNames } = readWorkbookPart(
        archive.get("xl/workbook.bin")!.data,
        "xl/workbook.bin"
      );
      const xlsx = Workbook.create();
      await Workbook.read(xlsx, await readFile(`${DIRECTORY}/${fixture}.excel.xlsx`));
      const compared: string[] = [];
      const differing: string[] = [];
      sheetNames.forEach((sheetName, sheet) => {
        const part = archive.get(`xl/worksheets/sheet${sheet + 1}.bin`);
        const worksheet = Workbook.getWorksheet(xlsx, sheetName);
        if (part === undefined || worksheet === undefined) {
          return;
        }
        for (const cell of formulaCells(part.data, sheet)) {
          const text = Cell.getFormula(worksheet, cell.address);
          if (text === undefined) {
            continue;
          }
          // The matrix's one table, as its `BrtBeginList` declares it — the package reader resolves this from the table
          // parts, which this test does not read.
          const context = {
            ...formulaContext,
            sheetIndex: sheet,
            tables: new Map([["Tbl", { id: 1, columns: ["Item", "Qty"], sheet: "Third" }]])
          };
          const where = `${sheetName}!${cell.address}`;
          // Decode: Excel's bytes are Excel's text.
          expect(
            formulaText(cell.rgce, { ...context, origin: cell }, where, cell.extra),
            where
          ).toBe(text);
          compared.push(text);
          const ours = encodeFormula(text, context, where);
          const same =
            maskUnused(ours.rgce, text.includes("{")) ===
              maskUnused(cell.rgce, text.includes("{")) &&
            [...ours.extra].join() === [...cell.extra].join();
          if (!same) {
            differing.push(text);
            // Whatever differs must still be the same formula.
            expect(formulaText(ours.rgce, context, where, ours.extra), where).toBe(text);
          }
        }
      });
      expect(compared.length).toBeGreaterThan(fixture === "formula-classes" ? 20 : 60);
      expect(differing.filter(text => !MEM_TOKEN_CHOICE.has(text))).toEqual([]);
    });
  }
);

/**
 * `LET` and `LAMBDA`, against the workbook Excel built by typing each formula (`data/let-oracle.excel.xlsb`, from the
 * `mk-oracle` probe): every formula is decoded with the workbook's own names and re-encoded, so a `PtgName` index means
 * the same name on both sides.
 *
 * The case that differed was `LET(sq,…,REDUCE(…))`: `REDUCE` as `LET`'s result is `22` in Excel's bytes, a reference,
 * where this wrote `42`. Excel then showed the formula with `@` and saved `_xlfn.SINGLE` around the call.
 */
describe("let-oracle: LET and LAMBDA tokens Excel wrote", () => {
  it("are what this encoder writes for the same text", async () => {
    const archive = await extractAll(await readFile(`${DIRECTORY}/let-oracle.excel.xlsb`));
    const { formulaContext, sheetNames } = readWorkbookPart(
      archive.get("xl/workbook.bin")!.data,
      "xl/workbook.bin"
    );
    const differing: string[] = [];
    let compared = 0;
    sheetNames.forEach((sheetName, sheet) => {
      const part = archive.get(`xl/worksheets/sheet${sheet + 1}.bin`);
      for (const cell of part === undefined ? [] : formulaCells(part.data, sheet)) {
        const context = { ...formulaContext, sheetIndex: sheet };
        const where = `${sheetName}!${cell.address}`;
        const text = formulaText(cell.rgce, { ...context, origin: cell }, where, cell.extra);
        if (typeof text !== "string") {
          continue;
        }
        compared++;
        const ours = encodeFormula(text, context, where);
        if (
          maskUnused(ours.rgce, false) !== maskUnused(cell.rgce, false) ||
          [...ours.extra].join() !== [...cell.extra].join()
        ) {
          differing.push(text);
        }
      }
    });
    expect(compared).toBe(12);
    expect(differing).toEqual([]);
  });
});
