import { collectMergeBorders, applyMergeBorders } from "@excel/utils/merge-borders";
import { describe, it, expect } from "vitest";

/**
 * Minimal cell stub used by the utility functions.
 */
function makeGrid(
  rows: number,
  cols: number,
  borders?: Record<string, any>
): Record<string, { style: Record<string, any> }> {
  const grid: Record<string, { style: Record<string, any> }> = {};
  for (let r = 1; r <= rows; r++) {
    for (let c = 1; c <= cols; c++) {
      grid[`${r},${c}`] = {
        style: borders ? { border: { ...borders } } : {}
      };
    }
  }
  return grid;
}

const find = (grid: Record<string, any>) => (r: number, c: number) => grid[`${r},${c}`];
const get = (grid: Record<string, any>) => (r: number, c: number) => {
  if (!grid[`${r},${c}`]) {
    grid[`${r},${c}`] = { style: {} };
  }
  return grid[`${r},${c}`];
};

describe("merge-borders", () => {
  describe("collectMergeBorders", () => {
    it("returns undefined when no cells have borders", () => {
      const grid = makeGrid(2, 2);
      const result = collectMergeBorders(1, 1, 2, 2, find(grid));
      expect(result).toBeUndefined();
    });

    it("returns undefined when cells exist but findCell returns undefined", () => {
      const result = collectMergeBorders(1, 1, 2, 2, () => undefined);
      expect(result).toBeUndefined();
    });

    it("collects top/bottom edges indexed by column offset", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(2, 3, { top: thin, bottom: thin, left: thin, right: thin });
      const result = collectMergeBorders(1, 1, 2, 3, find(grid));

      expect(result).toBeDefined();
      // topEdges: 3 columns
      expect(result!.topEdges).toHaveLength(3);
      expect(result!.topEdges[0]).toEqual(thin);
      expect(result!.topEdges[1]).toEqual(thin);
      expect(result!.topEdges[2]).toEqual(thin);
      // bottomEdges: 3 columns
      expect(result!.bottomEdges).toHaveLength(3);
      expect(result!.bottomEdges[0]).toEqual(thin);
    });

    it("collects left/right edges indexed by row offset", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(3, 2, { top: thin, bottom: thin, left: thin, right: thin });
      const result = collectMergeBorders(1, 1, 3, 2, find(grid));

      expect(result).toBeDefined();
      // leftEdges: 3 rows
      expect(result!.leftEdges).toHaveLength(3);
      expect(result!.leftEdges[0]).toEqual(thin);
      expect(result!.leftEdges[2]).toEqual(thin);
      // rightEdges: 3 rows
      expect(result!.rightEdges).toHaveLength(3);
      expect(result!.rightEdges[0]).toEqual(thin);
    });

    it("falls back to master border when slave has none", () => {
      const thin = { style: "thin" };
      const grid: Record<string, any> = {
        "1,1": { style: { border: { top: thin, left: thin, bottom: thin, right: thin } } },
        "1,2": { style: {} } // slave has no border
      };
      const result = collectMergeBorders(1, 1, 1, 2, find(grid));

      expect(result).toBeDefined();
      // Right edge at column 2 should fall back to master's right
      expect(result!.rightEdges[0]).toEqual(thin);
      // Top edge at column 2 should fall back to master's top
      expect(result!.topEdges[1]).toEqual(thin);
    });

    it("prefers slave border over master fallback", () => {
      const thin = { style: "thin" };
      const thick = { style: "thick" };
      const grid: Record<string, any> = {
        "1,1": { style: { border: { right: thin } } },
        "1,2": { style: { border: { right: thick } } }
      };
      const result = collectMergeBorders(1, 1, 1, 2, find(grid));

      expect(result).toBeDefined();
      // slave's own right takes precedence over master's right fallback
      expect(result!.rightEdges[0]).toEqual(thick);
    });

    it("captures diagonal and color from master", () => {
      const diag = { style: "thin", up: true, down: false };
      const color = { argb: "FFFF0000" };
      const grid: Record<string, any> = {
        "1,1": { style: { border: { diagonal: diag, color, top: { style: "thin" } } } },
        "1,2": { style: {} }
      };
      const result = collectMergeBorders(1, 1, 1, 2, find(grid));

      expect(result).toBeDefined();
      expect(result!.diagonal).toEqual(diag);
      expect(result!.color).toEqual(color);
    });

    it("returns collected borders even when only diagonal exists", () => {
      const grid: Record<string, any> = {
        "1,1": { style: { border: { diagonal: { style: "thin", up: true } } } },
        "1,2": { style: {} }
      };
      const result = collectMergeBorders(1, 1, 1, 2, find(grid));

      expect(result).toBeDefined();
      expect(result!.diagonal).toEqual({ style: "thin", up: true });
      // No edge borders
      expect(result!.topEdges.every(e => !e)).toBe(true);
    });

    it("handles single-row merge (top === bottom)", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(1, 3, { top: thin, bottom: thin, left: thin, right: thin });
      const result = collectMergeBorders(1, 1, 1, 3, find(grid));

      expect(result).toBeDefined();
      // Both top and bottom edges collected from the same row
      expect(result!.topEdges[0]).toEqual(thin);
      expect(result!.bottomEdges[0]).toEqual(thin);
      expect(result!.topEdges).toHaveLength(3);
      expect(result!.bottomEdges).toHaveLength(3);
    });

    it("handles single-column merge (left === right)", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(3, 1, { top: thin, bottom: thin, left: thin, right: thin });
      const result = collectMergeBorders(1, 1, 3, 1, find(grid));

      expect(result).toBeDefined();
      // Both left and right edges collected from the same column
      expect(result!.leftEdges[0]).toEqual(thin);
      expect(result!.rightEdges[0]).toEqual(thin);
      expect(result!.leftEdges).toHaveLength(3);
      expect(result!.rightEdges).toHaveLength(3);
    });
  });

  describe("applyMergeBorders", () => {
    it("assigns position-aware borders to each cell", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(2, 2, { top: thin, bottom: thin, left: thin, right: thin });
      const collected = collectMergeBorders(1, 1, 2, 2, find(grid))!;

      // Simulate merge: set all cells to master style
      for (let r = 1; r <= 2; r++) {
        for (let c = 1; c <= 2; c++) {
          grid[`${r},${c}`].style = { ...grid["1,1"].style };
        }
      }

      applyMergeBorders(1, 1, 2, 2, collected, get(grid));

      expect(grid["1,1"].style.border).toEqual({ left: thin, top: thin });
      expect(grid["1,2"].style.border).toEqual({ right: thin, top: thin });
      expect(grid["2,1"].style.border).toEqual({ left: thin, bottom: thin });
      expect(grid["2,2"].style.border).toEqual({ right: thin, bottom: thin });
    });

    it("removes border from interior cells", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(3, 3, { top: thin, bottom: thin, left: thin, right: thin });
      const collected = collectMergeBorders(1, 1, 3, 3, find(grid))!;

      // Simulate merge
      for (let r = 1; r <= 3; r++) {
        for (let c = 1; c <= 3; c++) {
          grid[`${r},${c}`].style = { ...grid["1,1"].style };
        }
      }

      applyMergeBorders(1, 1, 3, 3, collected, get(grid));

      // Interior cell (2,2) should have no border
      expect(grid["2,2"].style.border).toBeUndefined();
    });

    it("each cell gets an independent style copy", () => {
      const thin = { style: "thin" };
      const grid = makeGrid(1, 2, { top: thin, bottom: thin, left: thin, right: thin });
      const collected = collectMergeBorders(1, 1, 1, 2, find(grid))!;

      // Simulate merge
      grid["1,2"].style = { ...grid["1,1"].style };

      applyMergeBorders(1, 1, 1, 2, collected, get(grid));

      // Mutating one shouldn't affect the other
      grid["1,1"].style.border.top = { style: "double" };
      expect(grid["1,2"].style.border.top).toEqual(thin);
    });
  });
});
