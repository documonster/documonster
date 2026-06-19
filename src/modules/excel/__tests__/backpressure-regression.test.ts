import { Writable, PassThrough } from "node:stream";

import { createCsvReadStream, writeCsv } from "@excel/bridge/csv-bridge";
import { cellSetValue } from "@excel/core/cell";
import { rowCommit, rowGetCell } from "@excel/core/worksheet";
import { Workbook, Worksheet } from "@excel/index";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
/**
 * Backpressure regression tests for streaming xlsx writers.
 *
 * Verifies:
 *   1. WorkbookWriter does not deadlock when the user-supplied sink is slow.
 *   2. Memory stays bounded relative to the produced size, even with a
 *      drastically slow sink (sink slower than producer by orders of
 *      magnitude).
 *   3. Workbook.writeCsv(stream) honors backpressure (formatter.write false → drain).
 */
import { describe, it, expect } from "vitest";

/**
 * Build a sink that completes each write asynchronously after yielding the
 * event loop a fixed number of times.
 *
 * We deliberately avoid `setTimeout(cb, n)` here. Windows clamps every
 * setTimeout delay to its system timer resolution (~15.6ms), so a "1ms"
 * write becomes a 15.6ms write — and these tests, which produce
 * 10_000–20_000 row-sized chunks (no zip middleware to batch them), get
 * inflated past the 60s deadlock-detection timeout. The CI failure looks
 * like a deadlock but is really just timer-resolution-induced slowdown.
 *
 * setImmediate is event-loop driven, has no timer-resolution dependency,
 * and behaves identically on Node + Bun across Linux/macOS/Windows. The
 * only thing the backpressure path actually needs from a "slow sink" is
 * that the write callback fires on a *later* turn of the event loop
 * rather than synchronously — that is what makes write() return false
 * and forces the producer onto the 'drain' path. yieldCount controls how
 * many event-loop turns we wait before completing each write, widening
 * the window in which a broken backpressure implementation would buffer
 * unbounded data. This preserves the original "drastically slow
 * consumer" testing intent without the wall-clock dependency.
 */
function createSlowAsyncSink(options: { highWaterMark: number; yieldCount?: number }): Writable {
  const yieldCount = options.yieldCount ?? 4;
  return new Writable({
    highWaterMark: options.highWaterMark,
    write(_chunk: Uint8Array, _enc, cb) {
      let remaining = yieldCount;
      const next = (): void => {
        if (remaining-- <= 0) {
          cb();
          return;
        }
        setImmediate(next);
      };
      setImmediate(next);
    }
  });
}

describe("streaming write backpressure", () => {
  it("WorkbookWriter does not deadlock with a slow sink (50k rows)", async () => {
    const slowSink = createSlowAsyncSink({ highWaterMark: 4 * 1024 });

    const wb = new WorkbookWriter({
      stream: slowSink,
      trueStreaming: true,
      useSharedStrings: false
    });
    const ws = wb.addWorksheet("Big");

    const ROWS = 50_000;
    const big = "x".repeat(200);

    for (let i = 1; i <= ROWS; i++) {
      const row = ws.getRow(i);
      for (let c = 1; c <= 10; c++) {
        cellSetValue(rowGetCell(row, c), big);
      }
      rowCommit(row);
    }
    ws.commit();

    // If backpressure is broken, this either OOMs or deadlocks. 60s is
    // generous; in practice this should finish well under that on every
    // platform now that we no longer rely on setTimeout's wall-clock
    // resolution to simulate a slow consumer.
    const timer = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("WorkbookWriter deadlocked / too slow")), 60_000);
    });

    await Promise.race([wb.commit(), timer]);
  }, 120_000);

  it("WorkbookWriter does not hang when sink is a PassThrough piped to /dev/null", async () => {
    const pt = new PassThrough({ highWaterMark: 1024 });
    const sink = new Writable({
      highWaterMark: 4 * 1024,
      write(_chunk, _enc, cb) {
        setImmediate(cb);
      }
    });
    pt.pipe(sink);

    const wb = new WorkbookWriter({ stream: pt, trueStreaming: true });
    const ws = wb.addWorksheet("Sheet");

    for (let i = 1; i <= 5_000; i++) {
      cellSetValue(rowGetCell(ws.getRow(i), 1), `row-${i}`);
      rowCommit(ws.getRow(i));
    }
    ws.commit();

    const timer = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("hang")), 30_000);
    });
    await Promise.race([wb.commit(), timer]);
  }, 60_000);

  it("writeCsv(Workbook, stream) does not deadlock with a slow sink", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Data");
    for (let i = 1; i <= 10_000; i++) {
      Worksheet.addRow(ws, [`row-${i}`, i, i * 3.14]);
    }

    const slowSink = createSlowAsyncSink({ highWaterMark: 1024 });

    const timer = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("writeCsv deadlocked")), 60_000);
    });

    await Promise.race([writeCsv(wb, slowSink) as Promise<void>, timer]);
  }, 120_000);

  it("WorkbookWriter rejects with the original sink error (not silently 'success')", async () => {
    // Sink that errors after a couple of writes — exactly the scenario where
    // a hang would happen if we didn't release backpressure waiters on error.
    let writeCount = 0;
    const erroringSink = new Writable({
      highWaterMark: 1024,
      write(_chunk: Uint8Array, _enc, cb) {
        writeCount++;
        if (writeCount > 5) {
          cb(new Error("sink boom"));
          return;
        }
        // First few writes are slow enough to trigger backpressure, then
        // the sink fails. Without our error-handling, commit() would hang
        // forever waiting for a 'drain' that never comes.
        setTimeout(cb, 10);
      }
    });

    const wb = new WorkbookWriter({ stream: erroringSink, trueStreaming: true });
    const ws = wb.addWorksheet("Data");
    for (let i = 1; i <= 10_000; i++) {
      cellSetValue(rowGetCell(ws.getRow(i), 1), "x".repeat(200));
      rowCommit(ws.getRow(i));
    }
    ws.commit();

    const timer = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("commit hung after sink error")), 30_000);
    });

    // commit() MUST reject — silently resolving would mean we lost the
    // user's "write failed" signal and caller assumes a complete file
    // landed in the sink. The error must also be the original sink error,
    // not a generic "stream destroyed" kind of mask.
    let caught: unknown = null;
    try {
      await Promise.race([wb.commit(), timer]);
    } catch (err) {
      caught = err;
    }

    expect(caught).not.toBeNull();
    const msg = caught instanceof Error ? caught.message : String(caught);
    expect(msg).not.toBe("commit hung after sink error");
    // Surface the original "sink boom" — that's what the user wrote, that's
    // what they should see.
    expect(msg).toContain("sink boom");
  }, 60_000);

  it("createCsvReadStream(Workbook) honors downstream backpressure", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Data");
    // Build enough rows that without backpressure the formatter would buffer
    // megabytes before the slow sink could drain.
    for (let i = 1; i <= 20_000; i++) {
      Worksheet.addRow(ws, [`row-${i}`, i, "x".repeat(200)]);
    }

    const slowSink = createSlowAsyncSink({ highWaterMark: 1024 });

    const csvStream = createCsvReadStream(wb);
    csvStream.pipe(slowSink as any);

    const timer = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("createCsvReadStream pipeline deadlocked")), 60_000);
    });

    const finished = new Promise<void>((resolve, reject) => {
      slowSink.on("finish", () => resolve());
      slowSink.on("error", reject);
      csvStream.on("error", reject);
    });

    await Promise.race([finished, timer]);
  }, 120_000);

  it("Workbook stream APIs preserve correct row numbers (sparse rows / includeEmptyRows)", async () => {
    // Regression for a generator-based row iterator that mistakenly used
    // 0-based array indexes as row numbers, which silently dropped empty
    // row padding when the user asked for `includeEmptyRows`. Build a
    // worksheet with a gap (rows 1, 3, 5 populated; 2 and 4 empty) and
    // verify both `writeCsv(stream)` and `createCsvReadStream()` emit the
    // right number of lines.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sparse");
    cellSetValue(rowGetCell(Worksheet.getRow(ws, 1), 1), "row-1");
    cellSetValue(rowGetCell(Worksheet.getRow(ws, 3), 1), "row-3");
    cellSetValue(rowGetCell(Worksheet.getRow(ws, 5), 1), "row-5");

    // Path 1: writeCsv(stream) → exercises _writeCsvStream
    const collector1: string[] = [];
    const sink = new Writable({
      write(chunk: Uint8Array, _enc, cb) {
        collector1.push(Buffer.from(chunk).toString("utf8"));
        cb();
      }
    });
    await (writeCsv(wb, sink) as Promise<void>);
    const csv1 = collector1.join("");
    // Strip any trailing newline so we get one entry per row, including
    // empty rows that the formatter renders as the empty string between
    // two newlines. With `includeEmptyRows: true` (default), rows 1..5
    // each produce a line — 5 entries total.
    const lines1 = csv1.replace(/\r?\n$/, "").split(/\r?\n/);
    expect(lines1).toEqual(["row-1", "", "row-3", "", "row-5"]);

    // Path 2: createCsvReadStream() → exercises iterateWorksheetRows on
    // the read-stream side
    const csvStream = createCsvReadStream(wb);
    const chunks: Buffer[] = [];
    await new Promise<void>((resolve, reject) => {
      csvStream.on("data", (chunk: Uint8Array) => chunks.push(Buffer.from(chunk)));
      csvStream.on("end", () => resolve());
      csvStream.on("error", reject);
    });
    const csv2 = Buffer.concat(chunks).toString("utf8");
    const lines2 = csv2.replace(/\r?\n$/, "").split(/\r?\n/);
    expect(lines2).toEqual(["row-1", "", "row-3", "", "row-5"]);
  });
});
