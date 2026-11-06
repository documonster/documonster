/**
 * One read's options hold whichever container the bytes are and whichever entry point reads them: the `zip` limits,
 * `checkCrc32` and `maxRows` / `maxCols` used to be honoured by the buffered XLSX reader only — an XLSB read, a
 * diagnostic read and a streamed XLSX read accepted them and read the package anyway.
 */
import { ArchiveLimitExceededError } from "@archive/core/errors";
import { Cell, ExcelLimitExceededError, Workbook } from "@excel/index";
import { readXlsxStreamInto as readBrowserStream } from "@excel/xlsx/read/stream.browser";
import { describe, expect, it } from "vitest";

async function sample(
  format: "xlsx" | "xlsb",
  options: { level?: number } = {}
): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", "x".repeat(100));
  Cell.setValue(ws, "B1", 3);
  Cell.setValue(ws, "A2", 2);
  return Workbook.toBuffer(wb, {
    format,
    ...(options.level === undefined ? {} : { zip: { level: options.level } })
  });
}

async function failure(run: () => Promise<unknown>): Promise<unknown> {
  try {
    await run();
    return undefined;
  } catch (error) {
    return error;
  }
}

function expectLimit(error: unknown, limit: string, allowed: number): void {
  expect(error).toBeInstanceOf(ExcelLimitExceededError);
  expect(error).toMatchObject({ limit, allowed });
}

/** `bytes` with the first central-directory record's CRC-32 changed, so that part fails its check. */
function withBadCrc(bytes: Uint8Array): Uint8Array {
  const bad = bytes.slice();
  const view = new DataView(bad.buffer, bad.byteOffset, bad.byteLength);
  for (let i = 0; i < bad.length - 20; i++) {
    if (view.getUint32(i, true) === 0x02014b50) {
      view.setUint32(i + 16, (view.getUint32(i + 16, true) + 1) >>> 0, true);
      break;
    }
  }
  return bad;
}

describe("XLSB reads honour the read's options", () => {
  it.each([
    ["read", (bytes: Uint8Array, options: object) => Workbook.read(bytes, options)],
    [
      "readWithDiagnostics",
      (bytes: Uint8Array, options: object) => Workbook.readWithDiagnostics(bytes, options)
    ]
  ])("%s applies zip.maxEntries and zip.maxTotalUncompressedSize", async (_name, read) => {
    const bytes = await sample("xlsb");
    expectLimit(await failure(() => read(bytes, { zip: { maxEntries: 0 } })), "maxEntries", 0);
    const total = await failure(() => read(bytes, { zip: { maxTotalUncompressedSize: 10 } }));
    expectLimit(total, "maxTotalUncompressedSize", 10);
    expect((total as Error).cause).toBeInstanceOf(ArchiveLimitExceededError);
  });

  it("applies maxRows and maxCols as the XLSX reader does", async () => {
    for (const format of ["xlsx", "xlsb"] as const) {
      const bytes = await sample(format);
      expectLimit(await failure(() => Workbook.read(bytes, { maxRows: 1 })), "maxRows", 1);
      expectLimit(await failure(() => Workbook.read(bytes, { maxCols: 1 })), "maxCols", 1);
      await expect(Workbook.read(bytes, { maxRows: 2, maxCols: 2 })).resolves.toBeDefined();
    }
  });

  it("honours checkCrc32: false", async () => {
    const bad = withBadCrc(await sample("xlsb", { level: 0 }));
    await expect(Workbook.read(bad)).rejects.toMatchObject({ name: "ExcelFileError" });
    await expect(Workbook.read(bad, { zip: { checkCrc32: false } })).resolves.toBeDefined();
  });
});

describe("a streamed XLSX read holds the container limits", () => {
  it("counts parts and inflated bytes as they arrive", async () => {
    const bytes = await sample("xlsx");
    expectLimit(
      await failure(() => Workbook.readStream([bytes], { zip: { maxEntries: 0 } })),
      "maxEntries",
      0
    );
    expectLimit(
      await failure(() => Workbook.readStream([bytes], { zip: { maxEntrySize: 50 } })),
      "maxEntrySize",
      50
    );
    await expect(Workbook.readStream([bytes])).resolves.toBeDefined();
  });

  it("charges the bytes of a part it skips, as the buffered read does", async () => {
    const { Archive } = await import("@archive/index");
    const parts = await Archive.unpackZip(await sample("xlsx"));
    const entries: [string, Uint8Array][] = [...parts].map(([name, entry]) => [name, entry.data]);
    // A part the reader passes over without parsing: drained, and its bytes still count.
    entries.push([
      "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
      new Uint8Array(100_000).fill(32)
    ]);
    const bytes = Archive.packZipSync(entries, { level: 9 });
    expectLimit(
      await failure(() => Workbook.read(bytes, { zip: { maxEntrySize: 20_000 } })),
      "maxEntrySize",
      20_000
    );
    expectLimit(
      await failure(() => Workbook.readStream([bytes], { zip: { maxEntrySize: 20_000 } })),
      "maxEntrySize",
      20_000
    );
    await expect(Workbook.readStream([bytes])).resolves.toBeDefined();
  });
});

describe("the browser stream reader takes every input readStream declares", () => {
  it("reads a Web ReadableStream and an async iterable", async () => {
    const bytes = await sample("xlsx");
    const web = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(bytes);
        controller.close();
      }
    });
    const fromWeb = await readBrowserStream(Workbook.create(), web);
    expect(Cell.getValue(Workbook.getWorksheet(fromWeb, "S")!, "B1")).toBe(3);
    const fromIterable = await readBrowserStream(
      Workbook.create(),
      (async function* () {
        yield bytes;
      })()
    );
    expect(Cell.getValue(Workbook.getWorksheet(fromIterable, "S")!, "B1")).toBe(3);
  });

  it("decrypts with the read's password", async () => {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 7);
    const encrypted = await Workbook.toBuffer(wb, { password: "p", encryption: { spinCount: 1 } });
    const read = await readBrowserStream(Workbook.create(), [encrypted], { password: "p" });
    expect(Cell.getValue(Workbook.getWorksheet(read, "S")!, "A1")).toBe(7);
  });
});

describe("an XLSB write honours zip", () => {
  it("applies the compression level, and rebuilds an unchanged package that asks for one", async () => {
    const stored = await sample("xlsb", { level: 0 });
    const compressed = await sample("xlsb", { level: 9 });
    expect(stored.length).toBeGreaterThan(compressed.length);

    const unchanged = await Workbook.read(compressed);
    // Without zip settings an unchanged package comes back as it arrived; with them it is written as asked.
    expect(await Workbook.toBuffer(unchanged, { format: "xlsb" })).toEqual(compressed);
    const rewritten = await Workbook.toBuffer(unchanged, { format: "xlsb", zip: { level: 0 } });
    expect(rewritten.length).toBeGreaterThan(compressed.length);
  });
});
