/**
 * The streaming reader on a damaged package: every read settles, with the library's own error, and nothing reaches
 * the process as an uncaught exception after the caller has caught the failure.
 *
 * Both used to fail. A part whose XML broke before its CRC-32 was checked left the reader's pipe into the XML parser
 * holding an error nobody listened for, so the mismatch reported a moment later crashed the process. And a package
 * cut off inside a part left that part's stream open, so the reader waited for the rest of it forever.
 */

import { readFileSync, writeFileSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";

import { Row, Stream, Workbook } from "@excel/index";
import { BaseError } from "@utils/errors";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

const DATA = join(__dirname, "../../__tests__/data");

/** A stored (uncompressed) package whose `xl/workbook.xml` is malformed — and so no longer matches its CRC-32. */
async function brokenWorkbookPart(): Promise<Uint8Array> {
  const wb = Workbook.create();
  for (const name of ["A", "B"]) {
    const ws = Workbook.addWorksheet(wb, name);
    Row.addMany(
      ws,
      Array.from({ length: 50 }, (_, i) => [i, `row ${i}`])
    );
  }
  const xlsx = await Workbook.toBuffer(wb, { zip: { level: 0 } });
  const at = Buffer.from(xlsx).indexOf("<sheets>");
  expect(at).toBeGreaterThan(0);
  xlsx[at + 7] = "=".charCodeAt(0);
  return xlsx;
}

async function readAll(input: Uint8Array | string): Promise<number> {
  let rows = 0;
  for await (const ws of new Stream.WorkbookReader(input)) {
    for await (const row of ws) {
      void row;
      rows++;
    }
  }
  return rows;
}

function settleWithin<T>(promise: Promise<T>, ms: number): Promise<T> {
  return Promise.race([
    promise,
    new Promise<T>((_, reject) =>
      setTimeout(() => reject(new Error(`did not settle within ${ms} ms`)), ms)
    )
  ]);
}

const pause = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

let uncaught: unknown[] = [];
const onUncaught = (error: unknown): void => {
  uncaught.push(error);
};

beforeEach(() => {
  uncaught = [];
  process.on("uncaughtException", onUncaught);
});

afterEach(() => {
  process.off("uncaughtException", onUncaught);
});

describe("Stream.WorkbookReader on a damaged package", () => {
  it("rejects a malformed part without a late uncaught CRC-32 error (bytes)", async () => {
    const bad = await brokenWorkbookPart();
    await expect(settleWithin(readAll(bad), 5000)).rejects.toBeInstanceOf(BaseError);
    await pause(100);
    expect(uncaught).toEqual([]);
  });

  it("rejects a malformed part without a late uncaught CRC-32 error (file path)", async () => {
    const dir = mkdtempSync(join(tmpdir(), "dm-corrupt-"));
    try {
      const file = join(dir, "bad.xlsx");
      writeFileSync(file, await brokenWorkbookPart());
      await expect(settleWithin(readAll(file), 5000)).rejects.toBeInstanceOf(BaseError);
      await pause(100);
      expect(uncaught).toEqual([]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  for (const fixture of ["hyperlink-without-text.xlsx", "chart-pivot-sample.xlsx"]) {
    it(`settles on every truncation of ${fixture}`, async () => {
      const bytes = new Uint8Array(readFileSync(join(DATA, fixture)));
      const outcomes: string[] = [];
      for (let tenth = 1; tenth < 20; tenth++) {
        const cut = bytes.slice(0, Math.floor((bytes.length * tenth) / 20));
        const outcome = await settleWithin(readAll(cut), 5000).then(
          () => "completed",
          (error: unknown) =>
            error instanceof BaseError ? "typed" : `untyped ${String(error)} at ${tenth}/20`
        );
        outcomes.push(outcome);
      }
      expect(outcomes.filter(o => o !== "typed" && o !== "completed")).toEqual([]);
      await pause(50);
      expect(uncaught).toEqual([]);
    }, 120_000);
  }

  it("reports a part that is not UTF-8 as a malformed part, on every read path", async () => {
    const wb = Workbook.create();
    Row.addMany(Workbook.addWorksheet(wb, "S"), [["abc", "def"]]);
    const xlsx = await Workbook.toBuffer(wb, { zip: { level: 0 }, useSharedStrings: false });
    const at = Buffer.from(xlsx).indexOf("abc");
    expect(at).toBeGreaterThan(0);
    xlsx[at] = 0xff;
    for (const checkCrc32 of [true, false]) {
      const reads: (() => Promise<unknown>)[] = [
        async () => {
          for await (const ws of new Stream.WorkbookReader(xlsx, { checkCrc32 })) {
            for await (const row of ws) {
              void row;
            }
          }
        },
        () =>
          Workbook.readStream(Workbook.create(), Readable.from([xlsx]), { zip: { checkCrc32 } }),
        () => Workbook.read(Workbook.create(), xlsx, { zip: { checkCrc32 } })
      ];
      for (const read of reads) {
        const error = await settleWithin(read(), 5000).then(
          () => undefined,
          (e: unknown) => e
        );
        // A lenient path may read it with a replacement character; what it must not do is fail with a `TypeError`.
        if (error !== undefined) {
          expect(error).toBeInstanceOf(BaseError);
        }
      }
    }
    expect(uncaught).toEqual([]);
  });

  it("releases the package when the caller stops reading part-way", async () => {
    const bytes = new Uint8Array(readFileSync(join(DATA, "chart-pivot-sample.xlsx")));
    for await (const ws of new Stream.WorkbookReader(bytes)) {
      for await (const row of ws) {
        void row;
        break;
      }
      break;
    }
    await pause(100);
    expect(uncaught).toEqual([]);
  });
});
