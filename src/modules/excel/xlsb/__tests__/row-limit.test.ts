/**
 * A row header beyond the sheet's last row is refused, not modelled.
 *
 * `BrtRowHdr.rw` is at most 1,048,575 (MS-XLSB 2.4.726). A crafted file carried `0x44000000`, and the reader put a row
 * object at that index: the row array's length became 1,140,850,689, and every later walk over the rows — building the
 * worksheet model first of all — visited a billion holes. Reading a 14 KB file took a minute.
 */
import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import { XlsbParseError } from "@excel/errors";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { readWorksheetPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

async function sheetPart(): Promise<Uint8Array> {
  const handle = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(handle, "S"), "A2", 1);
  const archive = await extractAll(await Workbook.toBuffer(handle, { format: "xlsb" }));
  return archive.get("xl/worksheets/sheet1.bin")!.data.slice();
}

describe("BrtRowHdr beyond the grid", () => {
  it("reads the last row of the grid", async () => {
    const bytes = await sheetPart();
    for (const record of iterateInterpretableRecords(bytes, "s")) {
      if (recordSpec(record.id)?.name === "BrtRowHdr") {
        new DataView(record.payload.buffer, record.payload.byteOffset).setUint32(
          0,
          1_048_575,
          true
        );
      }
    }
    expect(() => readWorksheetPart(bytes, "xl/worksheets/sheet1.bin", [])).not.toThrow();
  });

  it("is a parse error past it", async () => {
    const bytes = await sheetPart();
    for (const record of iterateInterpretableRecords(bytes, "s")) {
      if (recordSpec(record.id)?.name === "BrtRowHdr") {
        new DataView(record.payload.buffer, record.payload.byteOffset).setUint32(
          0,
          0x44000000,
          true
        );
      }
    }
    expect(() => readWorksheetPart(bytes, "xl/worksheets/sheet1.bin", [])).toThrow(XlsbParseError);
  });
});
