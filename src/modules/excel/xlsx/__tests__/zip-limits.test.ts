/**
 * Container limits the XLSX reader applies by default (`XlsxReadOptions.zip`).
 */

import { ArchiveLimitError } from "@archive/core/errors";
import { extractAll } from "@archive/unzip/extract";
import { createZipSync } from "@archive/zip/zip-bytes";
import { Cell, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

async function sampleXlsx(): Promise<Uint8Array> {
  const handle = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(handle, "S"), 1, 1, 1);
  return new Uint8Array(await Workbook.toBuffer(handle, { format: "xlsx" }));
}

describe("xlsx container limits", () => {
  it("honours an explicit maxEntries", async () => {
    const bytes = await sampleXlsx();
    const err = await Workbook.read(Workbook.create(), bytes, { zip: { maxEntries: 3 } }).catch(
      e => e
    );
    expect(err).toBeInstanceOf(ArchiveLimitError);
    expect((err as ArchiveLimitError).limit).toBe("maxEntries");
  });

  it("rejects a package with more than 10,000 parts by default, and Infinity lifts it", async () => {
    const parts = await extractAll(await sampleXlsx());
    const entries = [...parts].map(([name, file]) => ({ name, data: file.data }));
    for (let i = 0; entries.length <= 10_000; i++) {
      entries.push({ name: `junk/${i}.bin`, data: new Uint8Array(0) });
    }
    const padded = createZipSync(entries, { level: 0 });

    const err = await Workbook.read(Workbook.create(), padded).catch(e => e);
    expect(err).toBeInstanceOf(ArchiveLimitError);
    expect((err as ArchiveLimitError).allowed).toBe(10_000);

    const wb = Workbook.create();
    await Workbook.read(wb, padded, { zip: { maxEntries: Infinity } });
    expect(Cell.getValue(Workbook.getWorksheet(wb, "S")!, "A1")).toBe(1);
  });

  it("honours an explicit maxTotalUncompressedSize", async () => {
    const err = await Workbook.read(Workbook.create(), await sampleXlsx(), {
      zip: { maxTotalUncompressedSize: 100 }
    }).catch(e => e);
    expect(err).toBeInstanceOf(ArchiveLimitError);
    expect((err as ArchiveLimitError).limit).toBe("maxTotalUncompressedSize");
  });
});
