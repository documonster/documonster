/**
 * A package part whose bytes no longer match the CRC-32 the package records for it is
 * reported, by the buffered and the streaming reader alike, instead of read as data.
 */
import { ZipCrc32MismatchError } from "@archive/core/errors";
import { Cell, Stream, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { describe, expect, it } from "vitest";

const MARKER = 987654321;

/** A stored (uncompressed) package, so one byte of a sheet's XML can be changed in place. */
async function storedWorkbook(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  for (let i = 1; i <= 50; i++) {
    Cell.setValue(ws, i, 1, i === 25 ? MARKER : i);
  }
  return Workbook.toBuffer(wb, { zip: { level: 0 } });
}

/** Turn the marker value into a different, still well-formed number. */
function corruptSheet(xlsx: Uint8Array): Uint8Array {
  const out = xlsx.slice();
  const at = new TextDecoder("latin1").decode(out).indexOf(String(MARKER));
  expect(at).toBeGreaterThan(0);
  out[at] = "8".charCodeAt(0);
  return out;
}

async function streamValues(input: Uint8Array, checkCrc32?: false): Promise<unknown[]> {
  const reader = new WorkbookReader(input, { worksheets: "emit", checkCrc32 });
  const values: unknown[] = [];
  for await (const ws of reader) {
    for await (const row of ws) {
      values.push(Stream.getCellValue(Stream.getRowCell(row, 1)));
    }
  }
  return values;
}

describe("XLSX package integrity", () => {
  it("Workbook.read rejects a damaged sheet part", async () => {
    const bad = corruptSheet(await storedWorkbook());
    await expect(Workbook.read(Workbook.create(), bad)).rejects.toThrow(damagedPackage());
    const lenient = await Workbook.read(Workbook.create(), bad, { zip: { checkCrc32: false } });
    expect(Cell.getValue(Workbook.getWorksheet(lenient, "S")!, 25, 1)).toBe(887654321);
  });

  it("the streaming reader rejects a damaged sheet part", async () => {
    const good = await storedWorkbook();
    expect((await streamValues(good))[24]).toBe(MARKER);
    const bad = corruptSheet(good);
    await expect(streamValues(bad)).rejects.toThrow(damagedPackage());
    expect((await streamValues(bad, false))[24]).toBe(887654321);
  });
});

/**
 * A damaged part is an `ExcelFileError` — which `isExcelError` answers for — with the archive's `ZipCrc32MismatchError` as
 * its cause, as a package that is not one is reported.
 */
function damagedPackage() {
  return expect.objectContaining({
    name: "ExcelFileError",
    cause: expect.any(ZipCrc32MismatchError)
  });
}
