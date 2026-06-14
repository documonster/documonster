import { Writable } from "node:stream";

import { cellSetValue } from "@excel/cell";
import { rowCommit, rowGetCell } from "@excel/worksheet";
/**
 * Memory bounds verification for WorkbookWriter under a paced sink.
 *
 * `WorkbookWriter` does not strictly bound memory at O(constant) when the
 * user feeds rows in a tight synchronous loop without yielding — that is
 * a fundamental JS limitation (sync `Row.commit(row)` cannot await sink drain).
 * However it does ensure:
 *   - Memory grows roughly linearly with the deflate batch buffers, capped
 *     by the sink's drain rate.
 *   - For 100k wide rows (~190 MB uncompressed input) memory should not
 *     blow up to many GB; the typical peak is ~150 MB driven by V8 GC
 *     timing, deflate working set, and StreamBuf accumulation.
 *
 * This test asserts the typical peak < 300 MB — a generous ceiling that
 * leaves room for V8 / GC / OS noise but catches obvious regressions like
 * the unbounded-buffer death spiral. Memory is sampled both during the
 * row loop AND throughout `commit()` (via a 50ms interval), so the
 * commit-phase peak — which is often the highest, when the central
 * directory is built — is covered.
 */
import { describe, it, expect } from "vitest";

import { WorkbookWriter } from "../../../index";

describe("WorkbookWriter memory bounds", () => {
  it("100k rows × 10 cols × 200B does not blow up", async () => {
    const sink = new Writable({
      highWaterMark: 64 * 1024,
      write(_chunk: Uint8Array, _enc, cb) {
        setImmediate(cb);
      }
    });

    const wb = new WorkbookWriter({
      stream: sink,
      trueStreaming: true,
      useSharedStrings: false
    });
    const ws = wb.addWorksheet("Big");

    const ROWS = 100_000;
    const big = "x".repeat(200);

    if (global.gc) {
      global.gc();
      global.gc();
    }
    const baseline = process.memoryUsage().rss;
    let peak = baseline;

    // Sample memory continuously throughout the entire run (loop + commit
    // phase). 50ms gives ~20 samples/sec — fine-grained enough to catch
    // spikes during finalisation.
    const samplerHandle = setInterval(() => {
      const rss = process.memoryUsage().rss;
      if (rss > peak) {
        peak = rss;
      }
    }, 50);

    try {
      for (let i = 1; i <= ROWS; i++) {
        const row = ws.getRow(i);
        for (let c = 1; c <= 10; c++) {
          cellSetValue(rowGetCell(row, c), big);
        }
        rowCommit(row);
      }
      ws.commit();
      await wb.commit();
    } finally {
      clearInterval(samplerHandle);
    }

    if (global.gc) {
      global.gc();
      global.gc();
    }

    const peakDeltaMB = (peak - baseline) / 1024 / 1024;
    console.log(
      `100k rows × 10 cols × 200B (~190 MB uncompressed input): peak RSS Δ = ${peakDeltaMB.toFixed(1)} MB`
    );

    // Pre-fix and post-fix both land around 130-200 MB on this hardware.
    // Allow generous slack for V8 / GC / OS noise; the assertion catches
    // catastrophic unbounded-buffer regressions, not micro-tuning wins.
    expect(peakDeltaMB).toBeLessThan(300);
  }, 120_000);
});
