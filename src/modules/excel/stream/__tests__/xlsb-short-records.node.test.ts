/**
 * A `.bin` record shorter than its fields is a malformed part. Every Excel read path reports it as the library's own
 * error — `XlsbParseError`, naming the part — instead of the `RangeError` the byte reader underneath raises, which a
 * caller cannot tell apart from a bug.
 */

import { Readable } from "node:stream";

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Row, Stream, Workbook, XlsbParseError } from "@excel/index";
import { encodeBiffRecords, iterateBiffRecords } from "@excel/xlsb/binary";
import { BaseError } from "@utils/errors";
import { describe, expect, it } from "vitest";

/** An `.xlsb` whose sheet records are cut short, framed correctly so the record walk itself still succeeds. */
async function xlsbWithShortRecords(ids: readonly number[], keep: number): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.addMany(ws, [
    [1, 2.5, "x"],
    [3, 4.5, "y"]
  ]);
  const xlsb = await Workbook.toBuffer(wb, { format: "xlsb" });
  const files = await extractAll(xlsb);
  const entries = [...files.values()]
    .filter(file => file.type === "file")
    .map(file => {
      if (!/^xl\/worksheets\/sheet\d+\.bin$/.test(file.path)) {
        return { name: file.path, data: file.data };
      }
      const records = [...iterateBiffRecords(file.data, file.path)].map(record => ({
        id: record.id,
        payload: ids.includes(record.id) ? record.payload.subarray(0, keep) : record.payload
      }));
      return { name: file.path, data: encodeBiffRecords(records) };
    });
  return createZip(entries);
}

async function streamRows(bytes: Uint8Array): Promise<void> {
  for await (const ws of new Stream.WorkbookReader(bytes)) {
    for await (const row of ws) {
      void row;
    }
  }
}

describe("XLSB records shorter than their fields", () => {
  it("fail the streaming reader with XlsbParseError (row header)", async () => {
    const bytes = await xlsbWithShortRecords([0x0000], 2);
    await expect(streamRows(bytes)).rejects.toBeInstanceOf(XlsbParseError);
  });

  it("fail the streaming reader with XlsbParseError (cells)", async () => {
    const bytes = await xlsbWithShortRecords([0x0002, 0x0005, 0x0007], 6);
    await expect(streamRows(bytes)).rejects.toBeInstanceOf(XlsbParseError);
  });

  for (const [label, ids, keep] of [
    ["row header", [0x0000], 2],
    ["cells", [0x0002, 0x0005, 0x0007], 6]
  ] as const) {
    it(`never surface a RangeError from Workbook.read / readStream (${label})`, async () => {
      const bytes = await xlsbWithShortRecords(ids, keep);
      for (const read of [
        () => Workbook.read(Workbook.create(), bytes, { format: "xlsb" }),
        () => Workbook.readStream(Workbook.create(), Readable.from([bytes]))
      ]) {
        const outcome = await read().then(
          () => undefined,
          (error: unknown) => error
        );
        if (outcome !== undefined) {
          expect(outcome).toBeInstanceOf(BaseError);
        }
      }
    });
  }
});
