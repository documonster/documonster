import { Writable } from "node:stream";

import { cellSetValue } from "@excel/core/cell";
import { rowCommit, rowGetCell } from "@excel/core/worksheet";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
/**
 * Large-streaming correctness under a paced (slow) sink.
 *
 * This test guards the streaming write path against the two failure modes a
 * large synchronous producer + slow consumer can trigger:
 *
 *   1. Deadlock / never-completing commit. If backpressure handling is broken,
 *      `commit()` can hang forever waiting for a drain that the pipeline never
 *      lets happen. The wall-clock timeout race catches that.
 *
 *   2. Data loss / corruption. If a pipeline stage drops or truncates buffered
 *      data under load, the produced file is incomplete. We read the whole
 *      workbook back and assert every dimension (row count, first/last cell
 *      values) survives the round-trip.
 *
 * Why NOT a heap/RSS ceiling:
 *   Previous revisions asserted a `heapUsed`/`rss` delta. That was fundamentally
 *   unmeasurable here. The suite runs with `isolate: false` (see
 *   vitest.config.ts), so ~400 test files share one V8 heap and one process
 *   RSS — any absolute reading is dominated by unrelated prior state. Worse,
 *   in the Node sync-deflate path NO compressed bytes reach the sink during the
 *   synchronous `row.commit()` loop (verified: sink receives 0 bytes until the
 *   async `commit()` phase), so the rows are necessarily held in memory until
 *   commit regardless of pipeline health — a heap reading cannot distinguish
 *   "correct, holding rows until commit" from "leaking". And without
 *   `--expose-gc` (which plain `vitest run` / `bun run vitest` do not set) a
 *   forced collection is unavailable, so natural GC timing swamps any delta:
 *   an intentional 200 MB leak was measured to move the delta by <30 MB and
 *   pass a 100 MB ceiling. Such a test is a green light that guards nothing.
 *
 *   Behavioural correctness — completes, and the bytes are all there and
 *   accurate — IS deterministic on Node, Bun, and any `isolate` setting, so
 *   that is what we assert.
 */
import { describe, it, expect } from "vitest";

describe("WorkbookWriter large streaming under a slow sink", () => {
  it("100k rows × 10 cols × 200B completes and round-trips intact", async () => {
    const chunks: Uint8Array[] = [];
    // Paced sink: completes each write on a later event-loop turn via
    // setImmediate, so the consumer is strictly slower than the synchronous
    // producer. `setImmediate` (not setTimeout) avoids Windows' ~15ms timer
    // clamp that would otherwise inflate this past the timeout.
    const sink = new Writable({
      highWaterMark: 64 * 1024,
      write(chunk: Uint8Array, _enc, cb) {
        chunks.push(chunk.slice());
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
    const COLS = 10;
    const big = "x".repeat(200);

    for (let i = 1; i <= ROWS; i++) {
      const row = ws.getRow(i);
      for (let c = 1; c <= COLS; c++) {
        cellSetValue(rowGetCell(row, c), big);
      }
      rowCommit(row);
    }
    ws.commit();

    // If backpressure is broken this either OOMs or hangs. 90s is generous;
    // in practice this finishes in a couple of seconds.
    const timer = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error("WorkbookWriter deadlocked / too slow")), 90_000);
    });
    await Promise.race([wb.commit(), timer]);

    // Round-trip: the produced file must be complete and accurate. A pipeline
    // that dropped or truncated buffered data under load would fail here.
    const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
    expect(totalLength).toBeGreaterThan(0);
    const xlsx = new Uint8Array(totalLength);
    let offset = 0;
    for (const chunk of chunks) {
      xlsx.set(chunk, offset);
      offset += chunk.length;
    }

    const readBack = Workbook.create();
    await Workbook.read(readBack, xlsx);
    const sheet = Workbook.getWorksheet(readBack, "Big")!;
    expect(Worksheet.rowCount(sheet)).toBe(ROWS);
    expect(Cell.getValue(sheet, "A1")).toBe(big);
    expect(Cell.getValue(sheet, `J${ROWS}`)).toBe(big);
  }, 120_000);
});
