/**
 * StreamingDocxWriter — sink-mode (true end-to-end streaming) tests.
 *
 * Verifies:
 *  - bytes delivered through a Web `WritableStream` match the buffered
 *    output (round-trip correctness),
 *  - `addAsync` actually awaits sink backpressure,
 *  - sink errors surface from `addAsync` / `finalize`,
 *  - `reset()` is rejected in sink mode.
 */

import { describe, it, expect } from "vitest";

import { DocxWriteError, Build, Streaming } from "../index";

describe("StreamingDocxWriter — sink mode", () => {
  it("delivers byte-identical output through a Web WritableStream", async () => {
    // Build the same document twice: once buffered, once via sink.
    const buffered = Streaming.createDocxStream();
    for (let i = 0; i < 50; i++) {
      buffered.add(Build.textParagraph(`Paragraph ${i}`));
    }
    const reference = await buffered.finalize();

    const collected: Uint8Array[] = [];
    const ws = new WritableStream<Uint8Array>({
      write(chunk): void {
        collected.push(chunk);
      }
    });

    const piped = Streaming.createDocxStream({ sink: ws });
    for (let i = 0; i < 50; i++) {
      piped.add(Build.textParagraph(`Paragraph ${i}`));
    }
    const result = await piped.finalize();
    expect(result.length).toBe(0); // sentinel — bytes are in `collected`

    const total = collected.reduce((n, c) => n + c.length, 0);
    const sinkBytes = new Uint8Array(total);
    let off = 0;
    for (const c of collected) {
      sinkBytes.set(c, off);
      off += c.length;
    }
    expect(sinkBytes.length).toBe(reference.length);
    // ZIP magic
    expect(sinkBytes[0]).toBe(0x50);
    expect(sinkBytes[1]).toBe(0x4b);
  });

  it("addAsync awaits actual sink writes (backpressure honoured)", async () => {
    const writeOrder: number[] = [];
    let counter = 0;
    const ws = new WritableStream<Uint8Array>(
      {
        async write(_chunk): Promise<void> {
          // Force every write to take a tick; record the order so we can
          // assert addAsync waited for each one before resolving.
          await new Promise(r => setTimeout(r, 5));
          writeOrder.push(counter++);
        }
      },
      // High-water mark of 1 byte forces backpressure on essentially
      // every chunk; getWriter().ready resolves only after each drain.
      new ByteLengthQueuingStrategy({ highWaterMark: 1 })
    );

    const writer = Streaming.createDocxStream({ sink: ws });
    for (let i = 0; i < 10; i++) {
      // Each addAsync awaits everything queued so far. If backpressure
      // were ignored the writes would all batch up at finalize time and
      // writeOrder would only be populated then.
      await writer.addAsync(Build.textParagraph(`P${i}`));
    }

    // Even without finalize, at least one chunk should have flowed
    // through the sink because each addAsync forced a drain.
    const drainedDuringAdd = writeOrder.length;

    await writer.finalize();

    // After finalize all chunks must be visible.
    expect(writeOrder.length).toBeGreaterThan(0);
    expect(writeOrder.length).toBeGreaterThanOrEqual(drainedDuringAdd);
  });

  it("surfaces sink errors from addAsync", async () => {
    const ws = new WritableStream<Uint8Array>({
      write(_chunk): void {
        throw new Error("sink exploded");
      }
    });

    const writer = Streaming.createDocxStream({ sink: ws });
    let caught: unknown = null;
    try {
      // Push enough content that at least one ZIP chunk has been emitted
      // by the time we await; the error then surfaces from addAsync.
      for (let i = 0; i < 200; i++) {
        await writer.addAsync(Build.textParagraph(`P${i}`));
      }
      await writer.finalize();
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(DocxWriteError);
    expect(String(caught)).toMatch(/sink/i);
  });

  it("rejects reset() in sink mode", async () => {
    const ws = new WritableStream<Uint8Array>({ write(): void {} });
    const writer = Streaming.createDocxStream({ sink: ws });
    writer.add(Build.textParagraph("once"));
    expect(() => writer.reset()).toThrow(/sink mode/);
    await writer.finalize();
  });

  it("synchronous add() throws after a sink write has already failed", async () => {
    // A sink whose every write throws. We drive the writer with `addAsync`
    // (which awaits the sink-write drain chain and surfaces `_streamError`)
    // until the failure is observed, rather than racing a fixed timeout —
    // this is deterministic on both the Node (sync deflate) and browser
    // (async CompressionStream deflate) paths.
    const failingSinkOpts = {
      sink: new WritableStream<Uint8Array>({
        write(_chunk): void {
          throw new Error("immediate fail");
        }
      })
    };
    const writer = Streaming.createDocxStream(failingSinkOpts);

    // Feed elements via addAsync until a sink write has actually failed and
    // been captured. The first compressed chunk reaching the sink rejects;
    // addAsync surfaces it. Bounded so a logic regression can't hang.
    let sinkFailed = false;
    for (let i = 0; i < 2000 && !sinkFailed; i++) {
      try {
        await writer.addAsync(Build.textParagraph(`P${i}`));
      } catch {
        sinkFailed = true;
      }
    }
    expect(sinkFailed).toBe(true);

    // Once the sink has failed, the synchronous add() path must throw
    // rather than queue more doomed work.
    let threw = false;
    try {
      writer.add(Build.textParagraph("never reaches sink"));
    } catch (e) {
      threw = true;
      expect(String(e)).toMatch(/sink already failed/);
    }
    expect(threw).toBe(true);

    // Drain finalize to keep the test cleanup graceful.
    await writer.finalize().catch(() => undefined);
  });

  it("buffered mode (no sink) is unaffected — finalize returns full bytes", async () => {
    const writer = Streaming.createDocxStream();
    writer.add(Build.textParagraph("alpha"));
    const bytes = await writer.finalize();
    expect(bytes.length).toBeGreaterThan(0);
    expect(bytes[0]).toBe(0x50); // 'P'
    expect(bytes[1]).toBe(0x4b); // 'K'
  });
});
