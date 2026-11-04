/**
 * Future functions, `LAMBDA` and structured references: the tokens this encoder writes against the tokens Excel wrote.
 *
 * `future-functions.xlsx` holds one call per future function (plus `LET`/`LAMBDA`, immediately invoked lambdas and
 * structured references) as this library writes it; `future-functions.excel.xlsb` is Excel 16's own save of that
 * workbook as XLSB. Every formula cell on sheet `F` is re-encoded from the XLSX text with Excel's own name table and
 * compared byte for byte, and Excel's bytes must read back to that text — as `operand-classes.node.test.ts` does for
 * the built-in functions.
 *
 * The spill-range operator is among them: the XLSX holds `SUM(_xlfn.ANCHORARRAY(D!H1))`, which reads back as
 * `SUM(D!H1#)` from either container and is written back to Excel's bytes for the call.
 *
 * What this pins: the operand class of each future function's token (`TRIMRANGE`, `ANCHORARRAY`, `TAKE` and `DROP`
 * return references — written in the value class Excel shows them as `@TRIMRANGE(…)`), the class of a call through a
 * `LET`-bound name, which arguments are arrays, `PtgAttrSemi` for a volatile one, `[#Data]` as `DATA2`, and the
 * encoding of an immediately invoked `LAMBDA`.
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
 * Cells where Excel's bytes say something other than the text, by address — each a normalisation Excel made on
 * reading the XLSX, so neither direction can match:
 *
 * - `A181`–`A184`: `ECMA.CEILING`, `ISO.CEILING`, `NETWORKDAYS.INTL` and `WORKDAY.INTL` carry no prefix in XLSX, and
 *   Excel's XLSB still calls them through `_xlfn.` stubs, so the stored name reads back prefixed.
 * - `A234` `+-+1`: Excel keeps only the unary plus in front, storing `+-1`.
 * - `A268` `SUM({-1,2.5E+1,1E-1})`: an array constant holds numbers, not their spelling.
 * - `A177` `_xlfn.SINGLE(D!A1:A3)`: a whole cell formula intersects its area anyway, so Excel writes the area alone
 *   in the value class and it reads back bare (see `bareIntersections` in `operand-classes.ts`).
 * - `A277` `SUM(Tb[@x]:Tb[@y])`: on a sheet without the table there is no "this row", and Excel stores both lists
 *   invalid. They read back as `#REF!:#REF!`, which is what Excel shows.
 */
const NORMALISED: ReadonlyMap<string, string> = new Map([
  ["A177", "D!A1:A3"],
  ["A181", "_xlfn.ECMA.CEILING(4.3,2)"],
  ["A182", "_xlfn.ISO.CEILING(4.3,2)"],
  ["A183", "_xlfn.NETWORKDAYS.INTL(45000,45030,11)"],
  ["A184", "_xlfn.WORKDAY.INTL(45000,10,11)"],
  ["A234", "+-1"],
  ["A268", "SUM({-1,25,0.1})"],
  ["A277", "SUM(#REF!:#REF!)"]
]);

interface ExcelCell {
  readonly address: string;
  readonly rgce: Uint8Array;
  readonly extra: Uint8Array;
}

function formulaCells(bytes: Uint8Array): ExcelCell[] {
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
    cells.push({ address: `${encodeCol(column)}${row + 1}`, rgce, extra });
  }
  return cells;
}

/**
 * The bytes the specification calls unused, which Excel fills with whatever was in memory: the two after a leading
 * `PtgAttrSemi`, and the last ten of `PtgArray`'s fourteen — the latter only where the text has an array constant, so
 * a `0x40` inside another token's payload masks nothing.
 */
function masked(rgce: Uint8Array, hasArray: boolean): string {
  const hex = [...rgce].map(byte => byte.toString(16).padStart(2, "0"));
  if (rgce[0] === 0x19 && rgce[1] === 0x01) {
    hex[2] = "..";
    hex[3] = "..";
  }
  for (let at = 0; hasArray && at < hex.length; at++) {
    if ((rgce[at]! === 0x20 || rgce[at]! === 0x40 || rgce[at]! === 0x60) && at + 15 <= hex.length) {
      for (let pad = at + 5; pad < at + 15; pad++) {
        hex[pad] = "..";
      }
    }
  }
  return hex.join(" ");
}

describe("future functions: tokens Excel wrote", () => {
  it("are what this encoder writes for the same text, and read back to that text", async () => {
    const archive = await extractAll(await readFile(`${DIRECTORY}/future-functions.excel.xlsb`));
    const { formulaContext, sheetNames } = readWorkbookPart(
      archive.get("xl/workbook.bin")!.data,
      "xl/workbook.bin"
    );
    const xlsx = Workbook.create();
    await Workbook.read(xlsx, await readFile(`${DIRECTORY}/future-functions.xlsx`));
    const sheet = sheetNames.indexOf("F");
    const worksheet = Workbook.getWorksheet(xlsx, "F")!;
    const context = {
      ...formulaContext,
      sheetIndex: sheet,
      tables: new Map([["Tb", { id: 1, columns: ["x", "y"], sheet: "Data" }]])
    };
    const differing: string[] = [];
    const misread: string[] = [];
    const texts: string[] = [];
    let compared = 0;
    for (const cell of formulaCells(archive.get(`xl/worksheets/sheet${sheet + 1}.bin`)!.data)) {
      const text = Cell.getFormula(worksheet, cell.address);
      if (text === undefined) {
        continue;
      }
      const where = `F!${cell.address}`;
      const read = formulaText(cell.rgce, context, where, cell.extra);
      if (read !== (NORMALISED.get(cell.address) ?? text)) {
        misread.push(`${where} ${text} read as ${read}`);
      }
      if (NORMALISED.has(cell.address)) {
        continue;
      }
      compared++;
      texts.push(text);
      const ours = encodeFormula(text, context, where);
      const hasArray = text.includes("{");
      if (
        masked(ours.rgce, hasArray) !== masked(cell.rgce, hasArray) ||
        [...ours.extra].join() !== [...cell.extra].join()
      ) {
        differing.push(
          `${where} ${text}\n  ours  ${masked(ours.rgce, hasArray)}\n  excel ${masked(cell.rgce, hasArray)}`
        );
      }
    }
    expect(compared).toBeGreaterThan(250);
    // The spill-range operator: Excel stored `SUM(_xlfn.ANCHORARRAY(D!H1))` with the `_xlfn.ANCHORARRAY` stub; both
    // readers give it back as `SUM(D!H1#)`, and this encoder writes that text to Excel's own bytes.
    expect(texts).toContain("SUM(D!H1#)");
    expect(misread).toEqual([]);
    expect(differing).toEqual([]);
  });
});

/** Every `BrtName` record of a workbook part, whole, by the name it defines. */
function nameRecords(workbookPart: Uint8Array): Map<string, string> {
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
    records.set(name, [...record.payload].map(byte => byte.toString(16).padStart(2, "0")).join(""));
  }
  return records;
}

describe("future-function, parameter and user-function names: records Excel wrote", () => {
  /**
   * The `_xlfn.`, `_xlpm.` and `_xludf.`-called names this writer creates, against Excel's own `BrtName` for the
   * same name: flags, body, and the four strings `fProc` brings — every byte. `_xlpm.` names written with flags 0
   * and no body were dropped by Excel on a re-save (`PtgName` index 0), losing the formulas that used them.
   */
  it.each([
    ["future-functions", /^_xl(fn|pm)\./],
    ["formula-edges", /^CONCAT$/]
  ])(
    "%s: every stub this writer emits is Excel's record byte for byte",
    async (fixture, pattern) => {
      const twin =
        fixture === "future-functions" ? "future-functions.xlsx" : `${fixture}.excel.xlsx`;
      const excel = nameRecords(
        (await extractAll(await readFile(`${DIRECTORY}/${fixture}.excel.xlsb`))).get(
          "xl/workbook.bin"
        )!.data
      );
      const workbook = Workbook.create();
      await Workbook.read(workbook, await readFile(`${DIRECTORY}/${twin}`));
      const ours = nameRecords(
        (
          await extractAll(
            await Workbook.toBuffer(workbook, { format: "xlsb", unsupported: "ignore" })
          )
        ).get("xl/workbook.bin")!.data
      );
      const compared = [...ours].filter(([name]) => pattern.test(name));
      expect(compared.length).toBeGreaterThan(0);
      const differing = compared
        .filter(([name, bytes]) => excel.get(name) !== bytes)
        .map(([name, bytes]) => `${name}\n  ours  ${bytes}\n  excel ${excel.get(name)}`);
      expect(differing).toEqual([]);
    }
  );
});
