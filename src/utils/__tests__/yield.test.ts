import {
  PACED_SLICE_LENGTH,
  createYieldBudget,
  paceChunks,
  runSteps,
  runStepsPaced,
  scheduleMacrotask,
  yieldToEventLoop
} from "@utils/yield";
import { afterEach, describe, expect, it, vi } from "vitest";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

async function collect<T>(chunks: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const chunk of chunks) {
    out.push(chunk);
  }
  return out;
}

describe("paceChunks", () => {
  const never = (): undefined => undefined;

  it("cuts a chunk into slices no longer than the slice length, keeping every byte", async () => {
    const big = new Uint8Array(PACED_SLICE_LENGTH * 2 + 5).map((_, i) => i % 251);
    const slices = await collect(paceChunks([big], never));
    expect(slices.map(slice => slice.length)).toEqual([PACED_SLICE_LENGTH, PACED_SLICE_LENGTH, 5]);
    const joined = new Uint8Array(big.length);
    let at = 0;
    for (const slice of slices) {
      joined.set(slice, at);
      at += slice.length;
    }
    expect(joined).toEqual(big);
  });

  it("hands a chunk within the slice length over as it is", async () => {
    const small = new Uint8Array([1, 2, 3]);
    const [only] = await collect(paceChunks([small], never));
    expect(only).toBe(small);
  });

  it("never cuts a string between the halves of a surrogate pair", async () => {
    const text = `${"a".repeat(PACED_SLICE_LENGTH - 1)}😀${"b".repeat(10)}`;
    const slices = await collect(paceChunks([text], never));
    expect(slices.join("")).toBe(text);
    for (const slice of slices) {
      expect(slice.isWellFormed()).toBe(true);
    }
  });

  it("asks the yield point after each slice, and waits for a turn it is given", async () => {
    const order: string[] = [];
    let asked = 0;
    const yieldPoint = (): Promise<void> | undefined => {
      asked++;
      return asked === 2
        ? new Promise(resolve => {
            order.push("yielding");
            scheduleMacrotask(() => {
              order.push("turn");
              resolve();
            });
          })
        : undefined;
    };
    for await (const slice of paceChunks([new Uint8Array(PACED_SLICE_LENGTH * 3)], yieldPoint)) {
      order.push(`slice ${slice.length}`);
    }
    expect(asked).toBe(3);
    expect(order).toEqual([
      `slice ${PACED_SLICE_LENGTH}`,
      `slice ${PACED_SLICE_LENGTH}`,
      "yielding",
      "turn",
      `slice ${PACED_SLICE_LENGTH}`
    ]);
  });
});

describe("createYieldBudget", () => {
  it("yields only once the budget has passed, and counts again from the yield", async () => {
    let clock = 1000;
    vi.spyOn(performance, "now").mockImplementation(() => clock);
    const yieldIfDue = createYieldBudget(10);

    clock += 9;
    expect(yieldIfDue()).toBeUndefined();
    clock += 1;
    const pending = yieldIfDue();
    expect(pending).toBeInstanceOf(Promise);
    await pending;

    clock += 9;
    expect(yieldIfDue()).toBeUndefined();
    clock += 1;
    expect(yieldIfDue()).toBeInstanceOf(Promise);
  });
});

describe("yieldToEventLoop", () => {
  it("runs after the microtasks and the macrotasks already queued", async () => {
    const order: string[] = [];
    scheduleMacrotask(() => order.push("macrotask"));
    queueMicrotask(() => order.push("microtask"));
    await yieldToEventLoop();
    order.push("resumed");
    expect(order).toEqual(["microtask", "macrotask", "resumed"]);
  });

  it("lets a timer fire during a loop that keeps yielding", async () => {
    let fired = false;
    setTimeout(() => (fired = true), 0);
    const deadline = Date.now() + 5000;
    while (!fired && Date.now() < deadline) {
      await yieldToEventLoop();
    }
    expect(fired).toBe(true);
  });
});

describe("the task each platform yields through", () => {
  it("setImmediate where there is one", async () => {
    const setImmediateSpy = vi.spyOn(globalThis, "setImmediate");
    await yieldToEventLoop();
    expect(setImmediateSpy).toHaveBeenCalledTimes(1);
  });

  it("not scheduler.yield, which a yielding loop would let starve timers", async () => {
    const schedulerYield = vi.fn(() => Promise.resolve());
    vi.stubGlobal("setImmediate", undefined);
    vi.stubGlobal("scheduler", { yield: schedulerYield });
    await yieldToEventLoop();
    expect(schedulerYield).not.toHaveBeenCalled();
  });

  it("a MessageChannel message in a browser that does not", async () => {
    vi.stubGlobal("setImmediate", undefined);
    const timeoutSpy = vi.spyOn(globalThis, "setTimeout");
    const order: string[] = [];
    scheduleMacrotask(() => order.push("first"));
    queueMicrotask(() => order.push("microtask"));
    await yieldToEventLoop();
    expect(order).toEqual(["microtask", "first"]);
    expect(timeoutSpy).not.toHaveBeenCalled();
  });

  it("setTimeout where nothing else exists", async () => {
    vi.stubGlobal("setImmediate", undefined);
    vi.stubGlobal("MessageChannel", undefined);
    const timeoutSpy = vi.spyOn(globalThis, "setTimeout");
    await yieldToEventLoop();
    expect(timeoutSpy).toHaveBeenCalledTimes(1);
  });
});

describe("runSteps / runStepsPaced", () => {
  function* counting(log: string[]): Generator<void, string, void> {
    for (let i = 0; i < 3; i++) {
      log.push(`step ${i}`);
      yield;
    }
    return "done";
  }

  it("runs every step and returns the result, synchronously", () => {
    const log: string[] = [];
    expect(runSteps(counting(log))).toBe("done");
    expect(log).toEqual(["step 0", "step 1", "step 2"]);
  });

  it("asks the yield point between the steps, and waits for a turn it is given", async () => {
    const log: string[] = [];
    let asked = 0;
    const yieldPoint = (): Promise<void> | undefined => {
      asked++;
      if (asked !== 2) {
        return undefined;
      }
      log.push("yielding");
      return yieldToEventLoop();
    };
    expect(await runStepsPaced(counting(log), yieldPoint)).toBe("done");
    expect(asked).toBe(3);
    expect(log).toEqual(["step 0", "step 1", "yielding", "step 2"]);
  });

  it("closes the steps when the yield point fails, so their finally runs", async () => {
    const log: string[] = [];
    function* restoring(): Generator<void, string, void> {
      log.push("set aside");
      try {
        yield;
        yield;
        return "done";
      } finally {
        log.push("restored");
      }
    }
    const failure = new Error("no turn");
    await expect(runStepsPaced(restoring(), () => Promise.reject(failure))).rejects.toBe(failure);
    expect(log).toEqual(["set aside", "restored"]);
  });

  it("runs that cleanup to its end even when it yields itself", async () => {
    const log: string[] = [];
    function* restoring(): Generator<void, string, void> {
      try {
        yield;
        return "done";
      } finally {
        log.push("cleanup started");
        yield;
        log.push("cleanup finished");
      }
    }
    await expect(
      runStepsPaced(restoring(), () => Promise.reject(new Error("no turn")))
    ).rejects.toThrow("no turn");
    expect(log).toEqual(["cleanup started", "cleanup finished"]);
  });
});
