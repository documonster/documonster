import { Readable } from "node:stream";

import { Crc32MismatchError } from "@archive/core/errors";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const MARKER = 987654321;

async function corruptStoredWorkbook(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  for (let i = 1; i <= 50; i++) {
    Cell.setValue(ws, i, 1, i === 25 ? MARKER : i);
  }
  const xlsx = await Workbook.toBuffer(wb, { zip: { level: 0 } });
  const at = Buffer.from(xlsx).indexOf(String(MARKER));
  xlsx[at] = "8".charCodeAt(0);
  return xlsx;
}

describe("XLSX package integrity (Node stream)", () => {
  it("Workbook.readStream rejects a damaged sheet part", async () => {
    const bad = await corruptStoredWorkbook();
    await expect(Workbook.readStream(Workbook.create(), Readable.from([bad]))).rejects.toThrow(
      damagedPackage()
    );
    const lenient = await Workbook.readStream(Workbook.create(), Readable.from([bad]), {
      zip: { checkCrc32: false }
    });
    expect(Cell.getValue(Workbook.getWorksheet(lenient, "S")!, 25, 1)).toBe(887654321);
  });
});

/**
 * A damaged part is an `ExcelFileError` — which `isExcelError` answers for — with the archive's `Crc32MismatchError` as
 * its cause, as a package that is not one is reported.
 */
function damagedPackage() {
  return expect.objectContaining({
    name: "ExcelFileError",
    cause: expect.any(Crc32MismatchError)
  });
}
