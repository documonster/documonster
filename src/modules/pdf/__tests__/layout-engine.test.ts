import { paginateRows } from "@pdf/render/layout-engine";
/**
 * Focused tests for layout engine pagination helpers.
 */
import { describe, expect, it } from "vitest";

describe("layout-engine pagination", () => {
  it("should repeat header rows on subsequent pages", () => {
    const pages = paginateRows([10, 10, 10, 10], 25, 1, new Set());
    expect(pages).toEqual([
      [0, 1],
      [0, 2],
      [0, 3]
    ]);
  });

  it("should avoid emitting repeat-row-only pages when headers cannot fit with body rows", () => {
    const pages = paginateRows([30, 30, 10], 35, 2, new Set());
    expect(pages).toEqual([[0], [1], [2]]);
  });

  it("should honor manual row breaks", () => {
    const pages = paginateRows([10, 10, 10, 10], 100, 0, new Set([2]));
    expect(pages).toEqual([
      [0, 1],
      [2, 3]
    ]);
  });
});
