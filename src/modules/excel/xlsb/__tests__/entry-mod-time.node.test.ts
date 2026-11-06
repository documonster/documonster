/**
 * Writing one workbook twice gives identical XLSB bytes, as it does for XLSX.
 *
 * Every ZIP entry carries a modification time. The XLSX writer stamps the workbook's own `modified ?? created`, so
 * its output depends only on the workbook; the XLSB writer used to stamp the moment of writing, so two writes a
 * couple of seconds apart (DOS times have two-second resolution) differed in every local header — a reproducible
 * build could not be had from the binary container at all. Both the buffered and the streamed path are covered,
 * because they reach the zip through different factories.
 */
import { Writable } from "node:stream";

import { Cell, Workbook } from "@excel";
import { afterEach, describe, expect, it, vi } from "vitest";

function build(): Workbook.Handle {
  const workbook = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(workbook, "S"), "A1", "alpha");
  return workbook;
}

async function streamed(workbook: Workbook.Handle): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const sink = new Writable({
    write(chunk: Uint8Array, _encoding, done) {
      chunks.push(chunk);
      done();
    }
  });
  await Workbook.writeStream(workbook, sink, { format: "xlsb" });
  return Buffer.concat(chunks);
}

describe("XLSB entry times come from the workbook, not the clock", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it.each([
    ["toBuffer", (wb: Workbook.Handle) => Workbook.toBuffer(wb, { format: "xlsb" })],
    ["writeStream", streamed]
  ] as const)("%s: two writes four seconds apart are byte-identical", async (_name, write) => {
    const workbook = build();
    vi.useFakeTimers({ toFake: ["Date"], now: new Date("2024-06-01T10:00:00Z") });
    const first = await write(workbook);
    vi.setSystemTime(new Date("2024-06-01T10:00:04Z"));
    const second = await write(workbook);
    expect(second).toEqual(first);
  });
});
