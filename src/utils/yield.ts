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
export function createYieldBudget(budgetMs = YIELD_BUDGET_MS): () => Promise<void> | undefined {
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
