import { createYieldBudget, scheduleMacrotask, yieldToEventLoop } from "@utils/yield";
import { afterEach, describe, expect, it, vi } from "vitest";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
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
