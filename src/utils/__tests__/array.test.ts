import { appendAll } from "@utils/array";
import { describe, expect, it } from "vitest";

describe("appendAll", () => {
  it("appends every item in order", () => {
    const target = [1, 2];
    appendAll(target, [3, 4, 5]);
    expect(target).toEqual([1, 2, 3, 4, 5]);
  });

  // What `push(...items)` cannot, on V8: each item an argument, the call stack overflows. (JavaScriptCore allows more,
  // so the native form is not asserted to fail.)
  it("appends more items than a call can take arguments", () => {
    const items = Array.from({ length: 500_000 }, (_, i) => i);
    const target: number[] = [];
    appendAll(target, items);
    expect(target).toHaveLength(500_000);
    expect(target[499_999]).toBe(499_999);
  });

  it("appends an array to itself once", () => {
    const items = [1, 2, 3];
    appendAll(items, items);
    expect(items).toEqual([1, 2, 3, 1, 2, 3]);
  });
});
