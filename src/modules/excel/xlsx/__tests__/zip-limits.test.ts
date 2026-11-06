/**
 * Container limits the XLSX reader applies by default (`XlsxReadOptions.zip`). A limit hit is an
 * `ExcelLimitExceededError` naming the option, with the archive's error as its cause.
 */

import { ArchiveLimitExceededError } from "@archive/core/errors";
import { extractAll } from "@archive/unzip/extract";
import { createZipSync } from "@archive/zip/zip-bytes";
import { Cell, ExcelLimitExceededError, Workbook, isExcelError } from "@excel";
import { describe, expect, it } from "vitest";

async function sampleXlsx(): Promise<Uint8Array> {
  const handle = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(handle, "S"), { row: 1, col: 1 }, 1);
  return new Uint8Array(await Workbook.toBuffer(handle, { format: "xlsx" }));
}

function expectLimit(err: unknown, limit: string, allowed?: number): void {
  expect(err).toBeInstanceOf(ExcelLimitExceededError);
  expect(isExcelError(err)).toBe(true);
  expect((err as ExcelLimitExceededError).limit).toBe(limit);
  if (allowed !== undefined) {
    expect((err as ExcelLimitExceededError).allowed).toBe(allowed);
  }
  expect((err as Error).cause).toBeInstanceOf(ArchiveLimitExceededError);
}

describe("xlsx container limits", () => {
  it("honours an explicit maxEntries", async () => {
    const bytes = await sampleXlsx();
    const err = await Workbook.read(Workbook.create(), bytes, { zip: { maxEntries: 3 } }).catch(
      e => e
    );
    expectLimit(err, "maxEntries", 3);
  });

  it("rejects a package with more than 10,000 parts by default, and Infinity lifts it", async () => {
    const parts = await extractAll(await sampleXlsx());
    const entries = [...parts].map(([name, file]) => ({ name, data: file.data }));
    for (let i = 0; entries.length <= 10_000; i++) {
      entries.push({ name: `junk/${i}.bin`, data: new Uint8Array(0) });
    }
    const padded = createZipSync(entries, { level: 0 });

    const err = await Workbook.read(Workbook.create(), padded).catch(e => e);
    expectLimit(err, "maxEntries", 10_000);

    const wb = Workbook.create();
    await Workbook.read(wb, padded, { zip: { maxEntries: Infinity } });
    expect(Cell.getValue(Workbook.getWorksheet(wb, "S")!, "A1")).toBe(1);
  });

  it("honours an explicit maxTotalUncompressedSize", async () => {
    const err = await Workbook.read(Workbook.create(), await sampleXlsx(), {
      zip: { maxTotalUncompressedSize: 100 }
    }).catch(e => e);
    expectLimit(err, "maxTotalUncompressedSize", 100);
  });
});
