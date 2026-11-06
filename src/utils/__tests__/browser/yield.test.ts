import { createYieldBudget, scheduleMacrotask, yieldToEventLoop } from "@utils/yield";
import { describe, expect, it } from "vitest";

describe("yieldToEventLoop (browser)", () => {
  it("does not need setImmediate", () => {
    expect((globalThis as { setImmediate?: unknown }).setImmediate).toBeUndefined();
  });

  it("runs a macrotask, after the queued microtasks", async () => {
    const order: string[] = [];
    queueMicrotask(() => order.push("microtask"));
    await new Promise<void>(resolve =>
      scheduleMacrotask(() => {
        order.push("macrotask");
        resolve();
      })
    );
    await yieldToEventLoop();
    expect(order).toEqual(["microtask", "macrotask"]);
  });

  it("lets a timer fire during a loop that keeps yielding", async () => {
    let fired = false;
    setTimeout(() => (fired = true), 0);
    const yieldIfDue = createYieldBudget(0);
    const deadline = Date.now() + 5000;
    while (!fired && Date.now() < deadline) {
      await yieldIfDue();
    }
    expect(fired).toBe(true);
  });
});
