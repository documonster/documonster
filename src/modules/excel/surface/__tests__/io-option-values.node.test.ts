/**
 * A compression failure inside the ZIP writer rejects the write instead of leaving its promise pending forever.
 *
 * Node-only: the failure is provoked with a level `zlib` refuses, reached through the internal writer below the
 * public option check. The browser's compressor clamps the level instead of failing, so there is nothing to observe
 * there.
 */
import { Cell, Workbook } from "@excel/index";
import { writeXlsxBytes } from "@excel/xlsx/write/package";
import { describe, expect, it } from "vitest";

function workbook() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  for (let i = 1; i <= 10; i++) {
    Cell.setValue(ws, `A${i}`, `v${i}`);
  }
  return wb;
}

function settlesWithin<T>(promise: Promise<T>, ms: number): Promise<"pending" | "settled"> {
  return Promise.race([
    promise.then(
      () => "settled" as const,
      () => "settled" as const
    ),
    new Promise<"pending">(resolve => setTimeout(() => resolve("pending"), ms))
  ]);
}

describe("ZIP writer failures", () => {
  it("a compression failure inside the ZIP writer rejects the write", async () => {
    const write = writeXlsxBytes(workbook(), { zip: { level: 42 } });
    expect(await settlesWithin(write, 3000)).toBe("settled");
    await expect(write).rejects.toThrow();
  });
});
