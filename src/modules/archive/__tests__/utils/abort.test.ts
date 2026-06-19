import {
  AbortError,
  createAbortError,
  createLinkedAbortController,
  isAbortError,
  throwIfAborted
} from "@archive/core/errors";
import { describe, expect, it } from "vitest";

describe("abort utils", () => {
  it("createAbortError should preserve reason as cause with fixed message", () => {
    const e0 = createAbortError();
    expect(e0).toBeInstanceOf(AbortError);
    expect(e0.name).toBe("AbortError");
    expect(e0.message).toBe("The operation was aborted");
    expect(e0.cause).toBeUndefined();

    const e1 = createAbortError("stop");
    expect(e1.message).toBe("The operation was aborted");
    expect(e1.cause).toBe("stop");

    const reason = new Error("oops");
    const e2 = createAbortError(reason);
    expect(e2.message).toBe("The operation was aborted");
    expect(e2.cause).toBe(reason);

    // Idempotent
    const e3 = createAbortError(e2);
    expect(e3).toBe(e2);
  });

  it("isAbortError should match AbortError-shaped errors", () => {
    expect(isAbortError(createAbortError("x"))).toBe(true);
    expect(isAbortError({ name: "AbortError" })).toBe(true);
    expect(isAbortError(new Error("x"))).toBe(false);
    expect(isAbortError(null)).toBe(false);
    expect(isAbortError(undefined)).toBe(false);
  });

  it("throwIfAborted should throw with signal.reason as cause and allow overriding reason", () => {
    const notAborted = new AbortController();
    expect(() => throwIfAborted(notAborted.signal)).not.toThrow();

    const ac = new AbortController();
    ac.abort("stop");

    try {
      throwIfAborted(ac.signal);
      expect.unreachable();
    } catch (e) {
      expect(isAbortError(e)).toBe(true);
      expect((e as any).cause).toBe("stop");
    }

    try {
      throwIfAborted(ac.signal, "override");
      expect.unreachable();
    } catch (e) {
      expect(isAbortError(e)).toBe(true);
      expect((e as any).cause).toBe("override");
    }
  });

  it("createLinkedAbortController should create independent controller when no parentSignal", () => {
    const { controller, cleanup } = createLinkedAbortController(undefined);

    expect(controller.signal.aborted).toBe(false);
    cleanup();
    expect(controller.signal.aborted).toBe(false);
  });

  it("createLinkedAbortController should abort immediately when parent already aborted", () => {
    const parent = new AbortController();
    parent.abort("stop");

    const { controller, cleanup } = createLinkedAbortController(parent.signal);

    expect(controller.signal.aborted).toBe(true);
    expect((controller.signal as any).reason).toBe("stop");

    // Cleanup should be safe / idempotent.
    cleanup();
    cleanup();
  });

  it("createLinkedAbortController should propagate abort reason and cleanup should detach", async () => {
    const parent = new AbortController();
    const { controller, cleanup } = createLinkedAbortController(parent.signal);

    expect(controller.signal.aborted).toBe(false);

    // Detach first; abort should NOT propagate.
    cleanup();
    parent.abort("stop");

    expect(controller.signal.aborted).toBe(false);

    // New link should propagate.
    const parent2 = new AbortController();
    const { controller: c2 } = createLinkedAbortController(parent2.signal);

    parent2.abort("stop2");
    expect(c2.signal.aborted).toBe(true);
    expect((c2.signal as any).reason).toBe("stop2");
  });
});
