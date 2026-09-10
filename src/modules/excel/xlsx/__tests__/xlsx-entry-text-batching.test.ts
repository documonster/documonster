import { extractAll } from "@archive/unzip/extract";
import { ZipDeflateFile } from "@archive/zip/stream";
import type { XlsxWritable } from "@excel/core/xlsx-io-types";
/**
 * Chunk granularity of the XLSX writer's ZIP adapter.
 *
 * `XmlStreamWriter` hands the adapter one string per XML tag — three per worksheet cell — and the adapter used
 * to answer each one with a `TextEncoder.encode()` and a push into the deflater. `encode()` costs far more per
 * call than per character, so a table-shaped workbook spent a fifth of its write time inside the encoder and
 * the deflater spent the rest concatenating thousands of scraps. The adapter now concatenates text until it is
 * worth encoding.
 *
 * These cases pin the properties that make it safe: pushes scale with bytes rather than with tags, buffered text
 * still leaves the entry ahead of any bytes written after it, a batch boundary never splits a UTF-8 character,
 * and a write after `end()` is refused rather than buffered into an entry that will never be flushed.
 */
import * as Workbook from "@excel/surface/workbook.browser";
import * as Worksheet from "@excel/surface/worksheet";
import type { IZipWriter, XlsxWriteOptions } from "@excel/xlsx/xlsx.browser";
import { XLSX } from "@excel/xlsx/xlsx.browser";
import { concatUint8Arrays } from "@utils/binary";
import { describe, expect, it, vi } from "vitest";

/** A sink that keeps every byte, so an entry can be read back out of the finished package. */
function createCollectingSink(chunks: Uint8Array[]): XlsxWritable {
  return {
    write(data) {
      // The zip pipeline only ever hands bytes to a sink; the string form of `XlsxWritable` is for callers.
      chunks.push(typeof data === "string" ? new TextEncoder().encode(data) : data);
      return true;
    },
    end() {},
    on() {
      return this;
    },
    once() {
      return this;
    },
    off() {
      return this;
    }
  };
}

/** Reaches the adapter directly: `createZipWriter` is where the writer would get it. */
class ExposedXlsx extends XLSX {
  zipWriter(options?: XlsxWriteOptions["zip"]): IZipWriter {
    return this.createZipWriter(options);
  }
}

async function finish(zip: IZipWriter, chunks: Uint8Array[]): Promise<Uint8Array> {
  await zip.waitForDrain();
  await new Promise<void>(resolve => {
    zip.on("finish", () => resolve());
    zip.finalize();
  });
  return concatUint8Arrays(chunks);
}

describe("XLSX entry text batching", () => {
  it("pushes once per batch of text, not once per XML tag", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    const rows = Array.from({ length: 400 }, (_, r) =>
      Array.from({ length: 20 }, (_, c) => r * 20 + c)
    );
    Worksheet.addAoa(sheet, rows, { origin: "A1" });

    const push = vi.spyOn(ZipDeflateFile.prototype, "push");
    try {
      const bytes = await new XLSX(workbook).writeBuffer({ validate: false });
      // Measured: 24,936 pushes before batching, 20 after. The bound is loose enough to survive a part being
      // added and still three orders of magnitude below the per-tag count it replaces.
      expect(push.mock.calls.length).toBeLessThan(200);
      expect(bytes.length).toBeGreaterThan(0);
    } finally {
      push.mockRestore();
    }
  });

  it("keeps buffered text ahead of bytes written after it", async () => {
    const chunks: Uint8Array[] = [];
    const zip = new ExposedXlsx(Workbook.create()).zipWriter({});
    zip.pipe(createCollectingSink(chunks));

    const entry = zip.createEntry("mixed.bin");
    // Short enough to stay in the text buffer, so the ordering depends on the flush rather than on the batch
    // happening to fill up. A binary part is why `createEntry` accepts bytes at all: encoding them as text
    // would turn every byte above 0x7F into two.
    entry.write("<a>");
    entry.write(new Uint8Array([0xff, 0x00, 0x80]));
    entry.write("</a>");
    entry.end();

    const entries = await extractAll(await finish(zip, chunks));
    const data = entries.get("mixed.bin")?.data;
    expect(data && Array.from(data)).toEqual([
      0x3c, 0x61, 0x3e, 0xff, 0x00, 0x80, 0x3c, 0x2f, 0x61, 0x3e
    ]);
  });

  it("splits text that exceeds the batch size and loses nothing", async () => {
    const chunks: Uint8Array[] = [];
    const zip = new ExposedXlsx(Workbook.create()).zipWriter({});
    zip.pipe(createCollectingSink(chunks));

    // Multi-byte on purpose: the batch is measured in characters, so a UTF-8 boundary must never fall inside
    // one — the whole point of concatenating before encoding rather than encoding into a fixed buffer.
    const piece = "中文 — é😀";
    const repeats = 20000;
    const entry = zip.createEntry("long.xml");
    for (let i = 0; i < repeats; i++) {
      entry.write(piece);
    }
    entry.end();

    const entries = await extractAll(await finish(zip, chunks));
    const text = new TextDecoder().decode(entries.get("long.xml")!.data);
    expect(text).toBe(piece.repeat(repeats));
  });

  it("refuses a write after end() instead of buffering it into the void", async () => {
    const chunks: Uint8Array[] = [];
    const zip = new ExposedXlsx(Workbook.create()).zipWriter({});
    zip.pipe(createCollectingSink(chunks));

    const entry = zip.createEntry("closed.xml");
    entry.write("<a/>");
    entry.end();

    // The short case is the one batching would hide: it never reaches the zip file, so without this guard the
    // bytes are dropped in silence. The long case would reach a finalized file and reject asynchronously.
    expect(() => entry.write("<b/>")).toThrow(/already closed/);
    expect(() => entry.write("x".repeat(70000))).toThrow(/already closed/);
    expect(() => entry.write(new Uint8Array([1, 2, 3]))).toThrow(/already closed/);
    // A redundant end() stays harmless, as it was before batching.
    expect(() => entry.end()).not.toThrow();

    const entries = await extractAll(await finish(zip, chunks));
    expect(new TextDecoder().decode(entries.get("closed.xml")!.data)).toBe("<a/>");
  });
});
