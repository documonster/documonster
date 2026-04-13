/**
 * Tests for shared-edge border resolution and border precedence.
 */
import { describe, expect, it } from "vitest";
import { borderPrecedence, resolveSharedBorders } from "@pdf/render/layout-engine";
import type { LayoutCell, LayoutBorder, PdfColor } from "@pdf/types";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const BLACK: PdfColor = { r: 0, g: 0, b: 0 };
const WHITE: PdfColor = { r: 1, g: 1, b: 1 };

function makeBorder(overrides: Partial<LayoutBorder> = {}): LayoutBorder {
  return {
    width: 0.25,
    color: BLACK,
    dashPattern: [],
    ...overrides
  };
}

function makeCell(overrides: Partial<LayoutCell> = {}): LayoutCell {
  return {
    text: "",
    rect: { x: 0, y: 0, width: 60, height: 15 },
    fontFamily: "Helvetica",
    fontSize: 11,
    bold: false,
    italic: false,
    strike: false,
    underline: false,
    textColor: BLACK,
    fillColor: null,
    horizontalAlign: "left",
    verticalAlign: "bottom",
    wrapText: false,
    borders: { top: null, right: null, bottom: null, left: null },
    borderInsets: { top: 0, right: 0, bottom: 0, left: 0 },
    colSpan: 1,
    rowSpan: 1,
    hyperlink: null,
    richText: null,
    indent: 0,
    textRotation: 0,
    textOverflowWidth: 0,
    ...overrides
  };
}

// ---------------------------------------------------------------------------
// borderPrecedence
// ---------------------------------------------------------------------------

describe("borderPrecedence", () => {
  it("thicker border wins", () => {
    const thin = makeBorder({ width: 0.25 });
    const thick = makeBorder({ width: 1 });
    expect(borderPrecedence(thick)).toBeGreaterThan(borderPrecedence(thin));
  });

  it("medium beats thin", () => {
    const thin = makeBorder({ width: 0.25 });
    const medium = makeBorder({ width: 0.5 });
    expect(borderPrecedence(medium)).toBeGreaterThan(borderPrecedence(thin));
  });

  it("solid beats dashed at same width", () => {
    const solid = makeBorder({ dashPattern: [] });
    const dashed = makeBorder({ dashPattern: [3, 2] });
    expect(borderPrecedence(solid)).toBeGreaterThan(borderPrecedence(dashed));
  });

  it("double beats single at same width and dash", () => {
    const single = makeBorder({ isDouble: false });
    const double = makeBorder({ isDouble: true });
    expect(borderPrecedence(double)).toBeGreaterThan(borderPrecedence(single));
  });

  it("darker color wins as tie-break", () => {
    const dark = makeBorder({ color: BLACK });
    const light = makeBorder({ color: WHITE });
    expect(borderPrecedence(dark)).toBeGreaterThan(borderPrecedence(light));
  });

  it("width dominates over style differences", () => {
    const thickDashed = makeBorder({ width: 1, dashPattern: [3, 2] });
    const thinSolid = makeBorder({ width: 0.25, dashPattern: [] });
    expect(borderPrecedence(thickDashed)).toBeGreaterThan(borderPrecedence(thinSolid));
  });
});

// ---------------------------------------------------------------------------
// resolveSharedBorders
// ---------------------------------------------------------------------------

describe("resolveSharedBorders", () => {
  it("should remove duplicate border on horizontal shared edge (equal borders)", () => {
    const thinRight = makeBorder();
    const thinLeft = makeBorder();
    const cellA = makeCell({ borders: { top: null, right: thinRight, bottom: null, left: null } });
    const cellB = makeCell({ borders: { top: null, right: null, bottom: null, left: thinLeft } });
    cellA.borderInsets.right = thinRight.width / 2;
    cellB.borderInsets.left = thinLeft.width / 2;

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);
    grid.set("0:1", cellB);

    resolveSharedBorders(grid, 1, 2);

    // One of them should be null, the other kept
    const aDraw = cellA.borders.right !== null;
    const bDraw = cellB.borders.left !== null;
    expect(aDraw).not.toBe(bDraw); // exactly one draws

    // Both insets should be non-zero (the line still exists for both)
    expect(cellA.borderInsets.right).toBeGreaterThan(0);
    expect(cellB.borderInsets.left).toBeGreaterThan(0);
  });

  it("should remove duplicate border on vertical shared edge (equal borders)", () => {
    const thinBottom = makeBorder();
    const thinTop = makeBorder();
    const cellA = makeCell({ borders: { top: null, right: null, bottom: thinBottom, left: null } });
    const cellB = makeCell({ borders: { top: thinTop, right: null, bottom: null, left: null } });
    cellA.borderInsets.bottom = thinBottom.width / 2;
    cellB.borderInsets.top = thinTop.width / 2;

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);
    grid.set("1:0", cellB);

    resolveSharedBorders(grid, 2, 1);

    const aDraw = cellA.borders.bottom !== null;
    const bDraw = cellB.borders.top !== null;
    expect(aDraw).not.toBe(bDraw);
    expect(cellA.borderInsets.bottom).toBeGreaterThan(0);
    expect(cellB.borderInsets.top).toBeGreaterThan(0);
  });

  it("should keep the thicker border on a shared edge", () => {
    const thick = makeBorder({ width: 1 });
    const thin = makeBorder({ width: 0.25 });
    const cellA = makeCell({ borders: { top: null, right: thick, bottom: null, left: null } });
    const cellB = makeCell({ borders: { top: null, right: null, bottom: null, left: thin } });
    cellA.borderInsets.right = thick.width / 2;
    cellB.borderInsets.left = thin.width / 2;

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);
    grid.set("0:1", cellB);

    resolveSharedBorders(grid, 1, 2);

    // cellA has thick → it wins
    expect(cellA.borders.right).toBe(thick);
    expect(cellB.borders.left).toBeNull();

    // Both insets reflect the thick border's half-width
    expect(cellA.borderInsets.right).toBe(0.5); // 1 / 2
    expect(cellB.borderInsets.left).toBe(0.5); // updated to winner's
  });

  it("should keep solid over dashed at same width", () => {
    const solid = makeBorder({ dashPattern: [] });
    const dashed = makeBorder({ dashPattern: [3, 2] });
    const cellA = makeCell({ borders: { top: null, right: dashed, bottom: null, left: null } });
    const cellB = makeCell({ borders: { top: null, right: null, bottom: null, left: solid } });
    cellA.borderInsets.right = dashed.width / 2;
    cellB.borderInsets.left = solid.width / 2;

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);
    grid.set("0:1", cellB);

    resolveSharedBorders(grid, 1, 2);

    // solid (cellB) wins
    expect(cellA.borders.right).toBeNull();
    expect(cellB.borders.left).toBe(solid);
  });

  it("should not touch borders that are not shared", () => {
    const topBorder = makeBorder();
    const bottomBorder = makeBorder();
    const cellA = makeCell({
      borders: { top: topBorder, right: null, bottom: bottomBorder, left: null }
    });

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);

    resolveSharedBorders(grid, 1, 1);

    // No neighbours → nothing changes
    expect(cellA.borders.top).toBe(topBorder);
    expect(cellA.borders.bottom).toBe(bottomBorder);
  });

  it("should not touch edges where only one side has a border", () => {
    const rightBorder = makeBorder();
    const cellA = makeCell({
      borders: { top: null, right: rightBorder, bottom: null, left: null }
    });
    const cellB = makeCell({ borders: { top: null, right: null, bottom: null, left: null } });

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);
    grid.set("0:1", cellB);

    resolveSharedBorders(grid, 1, 2);

    // cellA keeps its right border — no conflict
    expect(cellA.borders.right).toBe(rightBorder);
  });

  it("should handle a 2x2 grid correctly", () => {
    // All cells have all 4 borders (thin)
    const cells: LayoutCell[] = [];
    for (let i = 0; i < 4; i++) {
      cells.push(
        makeCell({
          borders: {
            top: makeBorder(),
            right: makeBorder(),
            bottom: makeBorder(),
            left: makeBorder()
          },
          borderInsets: { top: 0.125, right: 0.125, bottom: 0.125, left: 0.125 }
        })
      );
    }
    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cells[0]); // top-left
    grid.set("0:1", cells[1]); // top-right
    grid.set("1:0", cells[2]); // bottom-left
    grid.set("1:1", cells[3]); // bottom-right

    resolveSharedBorders(grid, 2, 2);

    // Shared vertical edge between col 0 and col 1 (row 0):
    // cells[0].right OR cells[1].left should be null, not both
    const h0 = (cells[0].borders.right !== null ? 1 : 0) + (cells[1].borders.left !== null ? 1 : 0);
    expect(h0).toBe(1);

    // Shared horizontal edge between row 0 and row 1 (col 0):
    const v0 = (cells[0].borders.bottom !== null ? 1 : 0) + (cells[2].borders.top !== null ? 1 : 0);
    expect(v0).toBe(1);

    // Shared vertical edge between col 0 and col 1 (row 1):
    const h1 = (cells[2].borders.right !== null ? 1 : 0) + (cells[3].borders.left !== null ? 1 : 0);
    expect(h1).toBe(1);

    // Shared horizontal edge between row 0 and row 1 (col 1):
    const v1 = (cells[1].borders.bottom !== null ? 1 : 0) + (cells[3].borders.top !== null ? 1 : 0);
    expect(v1).toBe(1);

    // Outer edges (not shared) should still be present
    expect(cells[0].borders.top).not.toBeNull();
    expect(cells[0].borders.left).not.toBeNull();
    expect(cells[1].borders.top).not.toBeNull();
    expect(cells[1].borders.right).not.toBeNull();
    expect(cells[2].borders.bottom).not.toBeNull();
    expect(cells[2].borders.left).not.toBeNull();
    expect(cells[3].borders.bottom).not.toBeNull();
    expect(cells[3].borders.right).not.toBeNull();

    // All insets should remain non-zero (line is still present)
    for (const c of cells) {
      expect(c.borderInsets.top).toBeGreaterThan(0);
      expect(c.borderInsets.right).toBeGreaterThan(0);
      expect(c.borderInsets.bottom).toBeGreaterThan(0);
      expect(c.borderInsets.left).toBeGreaterThan(0);
    }
  });

  it("should update loser inset to winner width on asymmetric conflict", () => {
    const thick = makeBorder({ width: 1 });
    const thin = makeBorder({ width: 0.25 });
    const cellA = makeCell({
      borders: { top: null, right: null, bottom: thick, left: null },
      borderInsets: { top: 0, right: 0, bottom: 0.5, left: 0 }
    });
    const cellB = makeCell({
      borders: { top: thin, right: null, bottom: null, left: null },
      borderInsets: { top: 0.125, right: 0, bottom: 0, left: 0 }
    });

    const grid = new Map<string, LayoutCell>();
    grid.set("0:0", cellA);
    grid.set("1:0", cellB);

    resolveSharedBorders(grid, 2, 1);

    // cellA's thick bottom wins
    expect(cellA.borders.bottom).toBe(thick);
    expect(cellB.borders.top).toBeNull();

    // cellB's top inset should now reflect the thick border
    expect(cellB.borderInsets.top).toBe(0.5); // 1 / 2
    // cellA's bottom inset stays at its original
    expect(cellA.borderInsets.bottom).toBe(0.5);
  });
});
