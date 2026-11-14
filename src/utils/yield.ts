/**
 * Handing control back to the event loop during long synchronous work.
 *
 * A loop that yielded with `setTimeout(0)` after every item spent most of its time idle: Node
 * waits at least a millisecond for a zero timeout, and a browser clamps nested timeouts to
 * four. Reading a 4,000-page PDF took about six seconds, mostly in those waits. So the work
 * yields only once it has run for a {@link YIELD_BUDGET_MS} budget, and through the cheapest
 * task the platform offers — which still lets timers, I/O, input and abort handlers run in
 * between, the reason for yielding at all.
 */

/**
 * A yield point, from {@link createYieldBudget}: call it between pieces of work and `await` what it returns — a turn of
 * the event loop once the budget has run out, `undefined` before.
 */
export type YieldPoint = () => Promise<void> | undefined;

/** How long a loop may run before {@link createYieldBudget} says it should yield. */
export const YIELD_BUDGET_MS = 10;

interface YieldGlobals {
  setImmediate?: (callback: () => void) => unknown;
  MessageChannel?: typeof MessageChannel;
}

/**
 * Run `callback` in a new macrotask, as soon as the event loop allows: `setImmediate` where it
 * exists (Node, Bun), otherwise a `MessageChannel` message (a browser or worker), and
 * `setTimeout(0)` only where neither does. Unlike a microtask, the event loop gets a turn first.
 */
export function scheduleMacrotask(callback: () => void): void {
  const g = globalThis as YieldGlobals;
  if (typeof g.setImmediate === "function") {
    g.setImmediate(callback);
    return;
  }
  if (typeof g.MessageChannel === "function") {
    // A channel per call, closed once used: an open port with a listener keeps some runtimes
    // alive, and a shared one would need its own queue.
    const { port1, port2 } = new g.MessageChannel();
    port1.onmessage = () => {
      port1.close();
      callback();
    };
    port2.postMessage(undefined);
    return;
  }
  setTimeout(callback, 0);
}

/**
 * Yield to the event loop through a {@link scheduleMacrotask}.
 *
 * Not `scheduler.yield()`, although a browser may have it: its continuation is queued ahead of
 * ordinary tasks, so a loop that keeps yielding through it never lets a timer fire — measured
 * in Chromium, where a `setTimeout(0)` was still pending after five seconds of such a loop. A
 * timer is how a caller bounds the work (`AbortSignal.timeout`), so it must get its turn.
 */
export function yieldToEventLoop(): Promise<void> {
  return new Promise(resolve => scheduleMacrotask(resolve));
}

function now(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

/**
 * A yield point for a long loop. Call the returned function between items: it yields to the
 * event loop once `budgetMs` of work has passed since the last yield (or since it was
 * created), and otherwise returns `undefined` — `await` it either way.
 */
export function createYieldBudget(budgetMs = YIELD_BUDGET_MS): YieldPoint {
  let last = now();
  return () => {
    if (now() - last < budgetMs) {
      return undefined;
    }
    return yieldToEventLoop().then(() => {
      last = now();
    });
  };
}

/** The most of a chunk {@link paceChunks} hands over at once: bytes, or UTF-16 code units of a string. */
export const PACED_SLICE_LENGTH = 64 * 1024;

/**
 * `chunks`, cut to at most {@link PACED_SLICE_LENGTH}, with `yieldPoint` (a {@link createYieldBudget}) asked after each
 * slice — so the event loop gets a turn once the work has run for its budget.
 *
 * For a streaming reader that parses each chunk synchronously and hands its results out through promises alone. On its
 * own such a reader lets the event loop run only when its source happens to wait — on a file read, or on zlib's thread
 * pool — and a source that already holds its data never waits: a stored (uncompressed) part of an archive in memory,
 * a part spooled in memory, a part inflated in one call. The whole part was then read in one task, however large. The
 * slice bounds the work between two yield points and the budget decides when to take one; both belong to the reader,
 * so neither depends on how the source chunks its data.
 */
export function paceChunks(
  chunks: AsyncIterable<Uint8Array> | Iterable<Uint8Array>,
  yieldPoint: YieldPoint
): AsyncGenerator<Uint8Array>;
export function paceChunks(
  chunks: AsyncIterable<string> | Iterable<string>,
  yieldPoint: YieldPoint
): AsyncGenerator<string>;
export function paceChunks(
  chunks: AsyncIterable<Uint8Array | string> | Iterable<Uint8Array | string>,
  yieldPoint: YieldPoint
): AsyncGenerator<Uint8Array | string>;
export async function* paceChunks(
  chunks: AsyncIterable<Uint8Array | string> | Iterable<Uint8Array | string>,
  yieldPoint: YieldPoint
): AsyncGenerator<Uint8Array | string> {
  for await (const chunk of chunks) {
    for (const slice of slices(chunk)) {
      yield slice;
      // Asked once the consumer wants more, so the time measured includes what it did with the slice.
      await yieldPoint();
    }
  }
}

/**
 * `chunk` in slices of at most {@link PACED_SLICE_LENGTH}, in order — the chunk itself when it is no longer than that.
 * A string is never cut between the two halves of a surrogate pair. A slice of bytes is a plain `Uint8Array` view,
 * whatever subclass the chunk was.
 */
function* slices(chunk: Uint8Array | string): Generator<Uint8Array | string> {
  if (chunk.length <= PACED_SLICE_LENGTH) {
    if (chunk.length > 0) {
      yield chunk;
    }
    return;
  }
  for (let start = 0; start < chunk.length;) {
    let end = Math.min(start + PACED_SLICE_LENGTH, chunk.length);
    if (typeof chunk === "string") {
      if (end < chunk.length) {
        const unit = chunk.charCodeAt(end - 1);
        if (unit >= 0xd800 && unit <= 0xdbff) {
          end--;
        }
      }
      yield chunk.slice(start, end);
    } else {
      yield new Uint8Array(chunk.buffer, chunk.byteOffset + start, end - start);
    }
    start = end;
  }
}

/**
 * Work written as a generator that yields between its steps: one implementation that a synchronous caller runs with
 * {@link runSteps} and an incremental one with {@link runStepsPaced}.
 */
export type Steps<R> = Generator<void, R, void>;

/** Run `steps` to the end, synchronously, and return what it returns. */
export function runSteps<R>(steps: Steps<R>): R {
  for (;;) {
    const step = steps.next();
    if (step.done) {
      return step.value;
    }
  }
}

/**
 * {@link runSteps}, asking `yieldPoint` (a {@link createYieldBudget}) between the steps.
 *
 * Steps left unfinished — the yield point failed — are closed and their cleanup run to its end, so a `finally` in them
 * completes: work written as steps may set state aside around them (a parse context it restores), and a generator
 * merely dropped never restores it. `return()` alone stops at a `yield` inside the `finally`, so the cleanup is driven
 * on — synchronously, the read being over.
 */
export async function runStepsPaced<R>(steps: Steps<R>, yieldPoint: YieldPoint): Promise<R> {
  let done = false;
  try {
    for (;;) {
      const step = steps.next();
      if (step.done) {
        done = true;
        return step.value;
      }
      await yieldPoint();
    }
  } finally {
    if (!done) {
      for (let step = steps.return(undefined as R); !step.done; step = steps.next()) {
        // cleanup still running
      }
    }
  }
}
