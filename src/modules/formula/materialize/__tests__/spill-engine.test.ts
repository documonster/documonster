/**
 * Unit tests for the persistent spill-engine WeakMap store.
 *
 * The spill engine's state lives across `calculateFormulas` invocations
 * so that ghost-cell detection and shrink cleanup can work correctly.
 * Its contract is narrow — "return a Map per workbook, reuse on
 * subsequent calls, allow GC when the workbook is gone" — but getting
 * it wrong leaks memory or drops ghost state between recalcs.
 */

import { describe, expect, it } from "vitest";

import { getGhostSnapshots, getPersistentSpillMap } from "../spill-engine";
import type { WorkbookLike } from "../types";

// The spill engine uses the workbook only as a WeakMap key; it never
// inspects `worksheets` / `getWorksheet`. Stub is cast to WorkbookLike
// rather than implementing the full interface — any real reference would
// work equivalently.
function stub(): WorkbookLike {
  return {} as WorkbookLike;
}

describe("getPersistentSpillMap", () => {
  it("returns a fresh Map for a new workbook", () => {
    const wb = stub();
    const m = getPersistentSpillMap(wb);
    expect(m).toBeInstanceOf(Map);
    expect(m.size).toBe(0);
  });

  it("returns the same Map on successive calls with the same workbook", () => {
    const wb = stub();
    const a = getPersistentSpillMap(wb);
    const b = getPersistentSpillMap(wb);
    expect(a).toBe(b);
  });

  it("gives distinct Maps for distinct workbooks", () => {
    const wb1 = stub();
    const wb2 = stub();
    expect(getPersistentSpillMap(wb1)).not.toBe(getPersistentSpillMap(wb2));
  });

  it("entries written through the returned Map persist", () => {
    const wb = stub();
    const m = getPersistentSpillMap(wb);
    m.set("ws:1!1:1", {
      worksheetId: 1,
      sourceRow: 1,
      sourceCol: 1,
      rows: 3,
      cols: 1
    });
    expect(getPersistentSpillMap(wb).size).toBe(1);
    expect(getPersistentSpillMap(wb).get("ws:1!1:1")?.rows).toBe(3);
  });
});

describe("getGhostSnapshots", () => {
  it("returns a fresh Map for a new workbook", () => {
    const wb = stub();
    const m = getGhostSnapshots(wb);
    expect(m).toBeInstanceOf(Map);
    expect(m.size).toBe(0);
  });

  it("returns the same Map on successive calls", () => {
    const wb = stub();
    const a = getGhostSnapshots(wb);
    const b = getGhostSnapshots(wb);
    expect(a).toBe(b);
  });

  it("ghost map and spill map are independent", () => {
    const wb = stub();
    const a = getPersistentSpillMap(wb);
    const b = getGhostSnapshots(wb);
    expect(a).not.toBe(b);
  });

  it("entries are typed as unknown — engine stores arbitrary ghost values", () => {
    const wb = stub();
    const m = getGhostSnapshots(wb);
    m.set("ws:1!2:1", 42);
    m.set("ws:1!3:1", "a string");
    m.set("ws:1!4:1", { error: "#N/A" });
    expect(m.size).toBe(3);
  });

  it("isolated per-workbook even when keys collide", () => {
    const wb1 = stub();
    const wb2 = stub();
    getGhostSnapshots(wb1).set("ws:1!2:1", 1);
    getGhostSnapshots(wb2).set("ws:1!2:1", 2);
    expect(getGhostSnapshots(wb1).get("ws:1!2:1")).toBe(1);
    expect(getGhostSnapshots(wb2).get("ws:1!2:1")).toBe(2);
  });
});

describe("spill-engine memory lifecycle", () => {
  it("does not retain workbook references via WeakMap", () => {
    // We can't directly force GC or observe retention in Vitest, but we
    // can at least verify the API's contract — calling with distinct
    // workbook objects doesn't leak entries into a shared bucket.
    const wb1 = stub();
    const wb2 = stub();
    const m1 = getPersistentSpillMap(wb1);
    const m2 = getPersistentSpillMap(wb2);
    m1.set("ws:1!1:1", { worksheetId: 1, sourceRow: 1, sourceCol: 1, rows: 1, cols: 1 });
    expect(m2.size).toBe(0);
  });
});
