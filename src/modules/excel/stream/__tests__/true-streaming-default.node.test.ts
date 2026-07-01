import { Writable } from "node:stream";

import { cellSetValue } from "@excel/core/cell";
import { rowCommit, rowGetCell } from "@excel/core/worksheet";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
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
 * Metric choice — `heapUsed`, not `rss`:
 *   The suite runs with `isolate: false` (see vitest.config.ts), so every
 *   test file shares one worker process and one V8 context. `rss` (resident
 *   set size) is a *process-wide* figure the OS almost never gives back once
 *   allocated — hundreds of prior test files leave the process RSS inflated
 *   and fragmented, which poisons any `rss` baseline taken here (observed:
 *   135 MB standalone vs 330+ MB inside the full suite, for identical
 *   producer behaviour). `heapUsed` measures live V8 heap objects and *is*
 *   reclaimed by `global.gc()`, so it isolates the memory this test actually
 *   retains regardless of what ran before. The regression this guards
 *   against — the unbounded-buffer death spiral — manifests as retained live
 *   heap, which `heapUsed` captures precisely.
 *
 * This test asserts the peak heap delta < 150 MB — a generous ceiling
 * (measured in-loop peak is ~35 MB) that leaves room for V8 / GC noise but
 * still catches the unbounded-buffer death spiral, which would retain
 * hundreds of MB of live heap.
 *
 * Sampling — in-loop manual + commit-phase interval:
 *   The 100k-row producer runs in a *tight synchronous loop* that never
 *   yields, so a `setInterval` sampler cannot fire during it (JS is
 *   single-threaded; timers only run at async boundaries). The loop is
 *   therefore where the peak occurs AND where an interval sampler is blind,
 *   so we sample `heapUsed` manually inside the loop. A `setInterval` sampler
 *   additionally covers the async `commit()` phase (central-directory build,
 *   final flushes), where control returns to the event loop.
 */
import { describe, it, expect } from "vitest";

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
    const baseline = process.memoryUsage().heapUsed;
    let peak = baseline;

    const sampleHeap = () => {
      const heapUsed = process.memoryUsage().heapUsed;
      if (heapUsed > peak) {
        peak = heapUsed;
      }
    };

    // The row loop below is fully synchronous and never yields, so a timer
    // callback cannot run during it. This interval only fires during the
    // async `commit()` phase — it's the commit-phase safety net. The row-loop
    // peak is captured by the manual `sampleHeap()` call inside the loop.
    const samplerHandle = setInterval(sampleHeap, 50);

    try {
      for (let i = 1; i <= ROWS; i++) {
        const row = ws.getRow(i);
        for (let c = 1; c <= 10; c++) {
          cellSetValue(rowGetCell(row, c), big);
        }
        rowCommit(row);
        // Sample inside the loop — this is where the producer peak actually
        // occurs and where the interval sampler is blind (see header comment).
        // Every 5000 rows keeps the overhead negligible while still catching
        // any monotonic buffer growth.
        if (i % 5000 === 0) {
          sampleHeap();
        }
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
      `100k rows × 10 cols × 200B (~190 MB uncompressed input): peak heap Δ = ${peakDeltaMB.toFixed(1)} MB`
    );

    // Measured in-loop peak is ~35 MB of retained live heap; 150 MB leaves
    // generous slack for V8 / GC noise while still catching a catastrophic
    // unbounded-buffer regression (which would retain hundreds of MB).
    expect(peakDeltaMB).toBeLessThan(150);
  }, 120_000);
});
