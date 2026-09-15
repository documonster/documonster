/**
 * Cutting a tall diagram across pages.
 *
 * The property that matters is not the piece count — it is *where* the cuts fall. A cut
 * through a line is fine and a cut through a box or a word is not, so these assert against the
 * geometry the display list actually contains rather than against coordinates, which would
 * have to be rewritten whenever a margin changed and would say nothing about whether the
 * picture survived.
 */

import { mermaidToDrawList } from "documonster/mermaid";
import { describe, expect, it } from "vitest";

import { planSlices } from "../tools/diagram-slice.js";

/** A chain of `n` boxes, one per rank: as tall as we like and narrow enough to need no fitting. */
const chain = (n: number): string =>
  `flowchart TB\n${Array.from(
    { length: n },
    (_, i) => `  S${i}["Step ${i + 1}"] --> S${i + 1}["Step ${i + 2}"]`
  ).join("\n")}`;

const listOf = (source: string) => mermaidToDrawList(source, {});

describe("planSlices", () => {
  it("leaves a diagram that fits in one piece", () => {
    const list = listOf(chain(2));
    const cuts = planSlices(list, 1, 600);
    expect(cuts).toHaveLength(1);
    expect(cuts[0]).toEqual({ top: 0, bottom: list.height, forced: false });
  });

  it("uses most of every page rather than cutting early", () => {
    // Greedy from the top, taking the *lowest* legal cut each time. Cutting at the first safe
    // row instead would be just as correct and would double the page count.
    for (const steps of [9, 16, 24]) {
      const list = listOf(chain(steps));
      const cuts = planSlices(list, 1, 600);
      for (const cut of cuts) {
        expect(cut.bottom - cut.top).toBeLessThanOrEqual(600);
      }
      // Every piece but the last fills at least three quarters of its page.
      for (const cut of cuts.slice(0, -1)) {
        expect(cut.bottom - cut.top).toBeGreaterThan(600 * 0.75);
      }
      expect(cuts.length).toBeLessThanOrEqual(Math.ceil(list.height / (600 * 0.75)));
    }
  });

  it("covers the whole diagram exactly once", () => {
    const list = listOf(chain(16));
    const cuts = planSlices(list, 1, 600);
    expect(cuts[0]!.top).toBe(0);
    expect(cuts.at(-1)!.bottom).toBeCloseTo(list.height, 5);
    for (let i = 1; i < cuts.length; i++) {
      expect(cuts[i]!.top).toBe(cuts[i - 1]!.bottom);
    }
  });

  it("never cuts through a box or a word", () => {
    // The whole point. Every node in this chain is a rect with text in it, so a cut inside one
    // would take half a box and half a label onto separate pages.
    const list = listOf(chain(16));
    const cuts = planSlices(list, 1, 600);
    const solids = solidBands(list);
    for (const cut of cuts.slice(0, -1)) {
      const through = solids.find(band => cut.bottom > band.top && cut.bottom < band.bottom);
      expect(
        through,
        `cut at ${cut.bottom} falls inside ${JSON.stringify(through)}`
      ).toBeUndefined();
      expect(cut.forced).toBe(false);
    }
  });

  it("treats a shape drawn as a filled polyline as solid", () => {
    // A rhombus, hexagon, parallelogram, stadium and cylinder are not rects — they are filled
    // polylines and paths, the same primitive a connector uses. Discriminating on `closed`
    // rather than on the fill let a cut fall straight through a diamond, so this pins the rule
    // by asking for a chain of nothing but diamonds.
    const diamonds = `flowchart TB\n${Array.from(
      { length: 14 },
      (_, i) => `  S${i}{"Decision ${i + 1}"} --> S${i + 1}{"Decision ${i + 2}"}`
    ).join("\n")}`;
    const list = listOf(diamonds);
    const cuts = planSlices(list, 1, 600);
    expect(cuts.length).toBeGreaterThan(1);
    const solids = solidBands(list).filter(band => band.what === "filled polyline");
    expect(solids.length).toBeGreaterThan(10);
    for (const cut of cuts.slice(0, -1)) {
      const through = solids.find(band => cut.bottom > band.top && cut.bottom < band.bottom);
      expect(through, `cut at ${cut.bottom} splits ${JSON.stringify(through)}`).toBeUndefined();
    }
  });

  it("ignores the theme's background plate", () => {
    // A theme paints one filled rect the size of the whole drawing. Counted as an obstacle it
    // covers every row, so nothing is cuttable and every cut is forced — and because the plan
    // still *looked* reasonable, this was invisible except in the notes. Both readings must
    // agree, which is the only way to state it: the backdrop changes no geometry.
    const source = chain(16);
    const plain = planSlices(mermaidToDrawList(source, {}), 1, 600);
    const themed = planSlices(
      mermaidToDrawList(source, { theme: { background: "#ffffff" } }),
      1,
      600
    );
    expect(themed.map(cut => [cut.top, cut.bottom, cut.forced])).toEqual(
      plain.map(cut => [cut.top, cut.bottom, cut.forced])
    );
    expect(themed.some(cut => cut.forced)).toBe(false);
  });

  it("scales the plan by the display fit, not the natural size", () => {
    // The budget is in points on the page; the bands are in the list's own units. Conflating
    // them cuts a diagram that was never going to overflow.
    const list = listOf(chain(16));
    expect(planSlices(list, 1, 600).length).toBeGreaterThan(1);
    expect(planSlices(list, 0.25, 600)).toHaveLength(1);
  });

  it("does not leave a postage-stamp final slice", () => {
    const list = {
      width: 100,
      height: 601,
      children: []
    };
    const cuts = planSlices(list, 1, 600);
    expect(cuts).toHaveLength(2);
    expect(cuts[0]!.bottom - cuts[0]!.top).toBeGreaterThanOrEqual(48);
    expect(cuts[1]!.bottom - cuts[1]!.top).toBeGreaterThanOrEqual(48);
  });

  it("does not cut through rotated text", () => {
    // A -90° label's vertical extent is its unrotated width. Treating it as
    // upright makes this look like a one-line obstacle around y=600 and chooses
    // a cut straight through the long rotated title.
    const list = {
      width: 200,
      height: 900,
      children: [
        {
          kind: "text" as const,
          x: 100,
          y: 600,
          rotate: -90,
          lines: [{ text: "A deliberately long vertical axis title", dy: 0 }],
          style: { size: 14, family: "Arial", anchor: "middle" as const }
        }
      ]
    };
    const first = planSlices(list, 1, 600)[0]!;
    expect(first.bottom).toBeLessThan(600);
    expect(first.forced).toBe(false);
  });

  it("says so when a box is taller than a page", () => {
    // Nothing can place this without a cut through the node. Reporting it is the difference
    // between a known compromise and a silent one.
    const tall = `flowchart TB\n  A["${"word ".repeat(400).trim()}"]`;
    const list = listOf(tall);
    const cuts = planSlices(list, 1, 300);
    expect(cuts.length).toBeGreaterThan(1);
    expect(cuts.some(cut => cut.forced)).toBe(true);
  });
});

/**
 * The vertical extents of everything solid, read back out of the display list.
 *
 * A second, independent reading of the same list: the planner walks it through the drawing
 * engine, and this recurses the tree by hand. Asserting a cut against the planner's own
 * measurement would only prove it is self-consistent.
 *
 * Groups are not transformed here, and do not need to be — a mermaid flowchart emits its marks
 * in the list's own coordinates. A producer that nested a transform would make this reading
 * wrong, and the test would fail rather than quietly pass.
 */
interface Solid {
  readonly top: number;
  readonly bottom: number;
  readonly what: string;
}

function solidBands(list: { readonly children: readonly unknown[] }): readonly Solid[] {
  const out: Solid[] = [];
  const walk = (children: readonly unknown[]): void => {
    for (const raw of children) {
      const node = raw as {
        kind: string;
        children?: readonly unknown[];
        paint?: { fill?: unknown };
        y?: number;
        height?: number;
        cy?: number;
        ry?: number;
        points?: readonly { y: number }[];
        style?: { size?: number };
        lines?: readonly unknown[];
      };
      switch (node.kind) {
        case "group":
          walk(node.children ?? []);
          break;
        case "rect":
          out.push({ top: node.y!, bottom: node.y! + node.height!, what: "rect" });
          break;
        case "ellipse":
          out.push({ top: node.cy! - node.ry!, bottom: node.cy! + node.ry!, what: "ellipse" });
          break;
        case "text": {
          const size = node.style?.size ?? 14;
          const count = Math.max(1, node.lines?.length ?? 1);
          out.push({
            top: node.y! - size,
            bottom: node.y! + (count - 1) * size * 1.2,
            what: "text"
          });
          break;
        }
        case "polyline":
          if (node.paint?.fill !== undefined && node.points !== undefined) {
            const ys = node.points.map(point => point.y);
            out.push({ top: Math.min(...ys), bottom: Math.max(...ys), what: "filled polyline" });
          }
          break;
        default:
          break;
      }
    }
  };
  walk(list.children);
  return out;
}
