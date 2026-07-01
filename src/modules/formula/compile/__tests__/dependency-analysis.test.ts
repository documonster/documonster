/**
 * Unit tests for the dependency analysis layer.
 *
 * Covers `buildDependencyGraphFromDeps`, `mergeDynamicDeps`, and
 * `topologicalSort`:
 * - Single-cell and range dependency expansion
 * - Cross-sheet dependency graphs
 * - Topological ordering respects dependencies
 * - Direct and indirect cycles are flagged via Tarjan's SCC
 * - Self-loops are detected
 * - Producer-map remapping (CSE / spill distribution)
 * - Dynamic dependency merging and its change-detection short-circuit
 */

import { describe, it, expect } from "vitest";

import type { DependencyGraph } from "../dependency-analysis";
import {
  buildDependencyGraphFromDeps,
  mergeDynamicDeps,
  topologicalSort
} from "../dependency-analysis";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Build the same cell key that the dependency module emits internally. */
function key(sheet: string, row: number, col: number): string {
  return `${sheet}!${row}:${col}`;
}

interface Cell {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
}

interface Area {
  readonly sheet: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

interface FormulaDescriptor {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  readonly deps?: readonly Cell[];
  readonly areas?: readonly Area[];
}

type CompiledMap = Map<
  string,
  {
    instance: { sheetName: string; row: number; col: number };
    staticDeps: { cells: readonly Cell[]; areas: readonly Area[] };
  }
>;

function buildCompiled(formulas: readonly FormulaDescriptor[]): CompiledMap {
  const map: CompiledMap = new Map();
  for (const f of formulas) {
    const k = key(f.sheet, f.row, f.col);
    map.set(k, {
      instance: { sheetName: f.sheet, row: f.row, col: f.col },
      staticDeps: {
        cells: f.deps ?? [],
        areas: f.areas ?? []
      }
    });
  }
  return map;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("dependency-analysis — single-cell dependencies", () => {
  it("records a forward edge from dependent formula to its source cell", () => {
    // B1 = A1. Only B1 is a formula.
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);

    const b1 = key("S", 1, 2);
    const a1 = key("S", 1, 1);

    expect(graph.formulaKeys).toEqual([b1]);
    expect(graph.dependsOn.get(b1)).toEqual(new Set([a1]));
    expect(graph.dependedBy.get(a1)).toEqual(new Set([b1]));
    expect(graph.circularKeys.size).toBe(0);
  });

  it("records a forward edge between two formula cells", () => {
    // A1 = (literal), B1 = A1, C1 = B1
    const compiled = buildCompiled([
      {
        sheet: "S",
        row: 1,
        col: 2,
        deps: [{ sheet: "S", row: 1, col: 1 }]
      },
      {
        sheet: "S",
        row: 1,
        col: 3,
        deps: [{ sheet: "S", row: 1, col: 2 }]
      }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);

    const b1 = key("S", 1, 2);
    const c1 = key("S", 1, 3);

    expect(graph.dependsOn.get(c1)).toEqual(new Set([b1]));
    expect(graph.dependedBy.get(b1)?.has(c1)).toBe(true);
  });

  it("builds no edges for a formula with no dependencies", () => {
    const compiled = buildCompiled([{ sheet: "S", row: 1, col: 1 }]);
    const graph = buildDependencyGraphFromDeps(compiled);

    const a1 = key("S", 1, 1);
    expect(graph.dependsOn.get(a1)?.size ?? 0).toBe(0);
  });
});

describe("dependency-analysis — range dependencies", () => {
  it("expands a small rectangular range into individual cell keys", () => {
    // C1 depends on the range A1:B2 — expected expansion: A1, A2, B1, B2.
    const compiled = buildCompiled([
      {
        sheet: "S",
        row: 1,
        col: 3,
        areas: [{ sheet: "S", top: 1, left: 1, bottom: 2, right: 2 }]
      }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);

    const c1 = key("S", 1, 3);
    const deps = graph.dependsOn.get(c1)!;
    expect(deps.size).toBe(4);
    expect(deps.has(key("S", 1, 1))).toBe(true);
    expect(deps.has(key("S", 1, 2))).toBe(true);
    expect(deps.has(key("S", 2, 1))).toBe(true);
    expect(deps.has(key("S", 2, 2))).toBe(true);
  });

  it("adds reverse edges for each expanded range cell", () => {
    const compiled = buildCompiled([
      {
        sheet: "S",
        row: 5,
        col: 1,
        areas: [{ sheet: "S", top: 1, left: 1, bottom: 1, right: 2 }]
      }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);

    const formula = key("S", 5, 1);
    expect(graph.dependedBy.get(key("S", 1, 1))?.has(formula)).toBe(true);
    expect(graph.dependedBy.get(key("S", 1, 2))?.has(formula)).toBe(true);
  });

  it("uses sparse scanning for a large range and only includes formula cells in-range", () => {
    // Dependent references the whole of column A (rows 1..1_048_576).
    // Only formula cells that actually fall inside the range should be
    // recorded as dependencies.
    const compiled = buildCompiled([
      // Formula cells inside the column:
      { sheet: "S", row: 10, col: 1 },
      { sheet: "S", row: 100, col: 1 },
      // Formula cell outside the column:
      { sheet: "S", row: 5, col: 2 },
      // Dependent formula.
      {
        sheet: "S",
        row: 1,
        col: 5,
        areas: [{ sheet: "S", top: 1, left: 1, bottom: 1_048_576, right: 1 }]
      }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);

    const dep = key("S", 1, 5);
    const deps = graph.dependsOn.get(dep)!;
    expect(deps.has(key("S", 10, 1))).toBe(true);
    expect(deps.has(key("S", 100, 1))).toBe(true);
    expect(deps.has(key("S", 5, 2))).toBe(false);
  });

  it("tracks cross-sheet range dependencies", () => {
    const compiled = buildCompiled([
      {
        sheet: "S1",
        row: 1,
        col: 1,
        areas: [{ sheet: "S2", top: 1, left: 1, bottom: 1, right: 2 }]
      }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    const dep = key("S1", 1, 1);
    expect(graph.dependsOn.get(dep)?.has(key("S2", 1, 1))).toBe(true);
    expect(graph.dependsOn.get(dep)?.has(key("S2", 1, 2))).toBe(true);
  });
});

describe("dependency-analysis — topological sort", () => {
  it("orders a simple chain A → B → C", () => {
    // A1 independent, B1 = A1, C1 = B1
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1 },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] },
      { sheet: "S", row: 1, col: 3, deps: [{ sheet: "S", row: 1, col: 2 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    const order = topologicalSort(graph);

    const a = order.indexOf(key("S", 1, 1));
    const b = order.indexOf(key("S", 1, 2));
    const c = order.indexOf(key("S", 1, 3));
    expect(a).toBeLessThan(b);
    expect(b).toBeLessThan(c);
    expect(order).toHaveLength(3);
  });

  it("orders a diamond so both mid nodes precede the sink", () => {
    // A → B, A → C, B → D, C → D
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1 },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] },
      { sheet: "S", row: 1, col: 3, deps: [{ sheet: "S", row: 1, col: 1 }] },
      {
        sheet: "S",
        row: 1,
        col: 4,
        deps: [
          { sheet: "S", row: 1, col: 2 },
          { sheet: "S", row: 1, col: 3 }
        ]
      }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    const order = topologicalSort(graph);

    const a = order.indexOf(key("S", 1, 1));
    const b = order.indexOf(key("S", 1, 2));
    const c = order.indexOf(key("S", 1, 3));
    const d = order.indexOf(key("S", 1, 4));
    expect(a).toBeLessThan(b);
    expect(a).toBeLessThan(c);
    expect(b).toBeLessThan(d);
    expect(c).toBeLessThan(d);
  });

  it("appends circular-reference nodes at the end in original order", () => {
    // A → B → A (cycle). C is independent.
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1, deps: [{ sheet: "S", row: 1, col: 2 }] },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] },
      { sheet: "S", row: 1, col: 3 }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    const order = topologicalSort(graph);

    const a = key("S", 1, 1);
    const b = key("S", 1, 2);
    const c = key("S", 1, 3);

    // C (independent) is ordered before the circular pair.
    expect(order.indexOf(c)).toBeLessThan(order.indexOf(a));
    expect(order.indexOf(c)).toBeLessThan(order.indexOf(b));
    // Circular entries appear in original insertion order.
    expect(order.indexOf(a)).toBeLessThan(order.indexOf(b));
  });
});

describe("dependency-analysis — circular reference detection", () => {
  it("detects a direct two-node cycle A ↔ B", () => {
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1, deps: [{ sheet: "S", row: 1, col: 2 }] },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    expect(graph.circularKeys.has(key("S", 1, 1))).toBe(true);
    expect(graph.circularKeys.has(key("S", 1, 2))).toBe(true);
  });

  it("detects a self-loop A = A", () => {
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    expect(graph.circularKeys.has(key("S", 1, 1))).toBe(true);
  });

  it("detects a three-node cycle A → B → C → A", () => {
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1, deps: [{ sheet: "S", row: 1, col: 2 }] },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 3 }] },
      { sheet: "S", row: 1, col: 3, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    expect(graph.circularKeys.size).toBe(3);
  });

  it("does not flag a DAG as circular", () => {
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1 },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    expect(graph.circularKeys.size).toBe(0);
  });

  it("flags every node in a cycle even when additional DAG edges feed it", () => {
    // Cycle A → B → A; plus independent D → A.
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1, deps: [{ sheet: "S", row: 1, col: 2 }] },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] },
      { sheet: "S", row: 1, col: 4, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    expect(graph.circularKeys.has(key("S", 1, 1))).toBe(true);
    expect(graph.circularKeys.has(key("S", 1, 2))).toBe(true);
    // D is NOT part of the cycle, only observes it.
    expect(graph.circularKeys.has(key("S", 1, 4))).toBe(false);
  });
});

describe("dependency-analysis — producer map (CSE/spill)", () => {
  it("redirects a dependency away from a non-formula cell to its producer", () => {
    // Setup: formula at B1 produces a spill into C1 (not itself a formula).
    // D1 references C1 — after remapping, D1 depends on the producer B1.
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 2 }, // B1 — producer (no deps)
      { sheet: "S", row: 1, col: 4, deps: [{ sheet: "S", row: 1, col: 3 }] } // D1 → C1
    ]);
    const producerMap = new Map<string, string>([[key("S", 1, 3), key("S", 1, 2)]]);
    const graph = buildDependencyGraphFromDeps(compiled, producerMap);

    const d = key("S", 1, 4);
    expect(graph.dependsOn.get(d)?.has(key("S", 1, 2))).toBe(true);
    expect(graph.dependsOn.get(d)?.has(key("S", 1, 3))).toBe(false);
  });
});

describe("dependency-analysis — mergeDynamicDeps", () => {
  function buildSimpleGraph(): DependencyGraph {
    // A1 (independent), B1 = A1
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1 },
      { sheet: "S", row: 1, col: 2, deps: [{ sheet: "S", row: 1, col: 1 }] }
    ]);
    return buildDependencyGraphFromDeps(compiled);
  }

  it("returns the same graph and changed=false when dynamicDeps is empty", () => {
    const graph = buildSimpleGraph();
    const out = mergeDynamicDeps(graph, new Map());
    expect(out.changed).toBe(false);
    expect(out.graph).toBe(graph);
  });

  it("returns the same graph when no new edges would be added", () => {
    const graph = buildSimpleGraph();
    // Same edge B1 → A1 as already-present in the static graph.
    const dynamic = new Map<string, Set<string>>();
    dynamic.set(key("S", 1, 2), new Set([key("S", 1, 1)]));

    const out = mergeDynamicDeps(graph, dynamic);
    expect(out.changed).toBe(false);
    expect(out.graph).toBe(graph);
  });

  it("adds a new dynamic edge and updates both forward and reverse indices", () => {
    const graph = buildSimpleGraph();
    // B1 dynamically accesses Z99.
    const dynamic = new Map<string, Set<string>>();
    dynamic.set(key("S", 1, 2), new Set([key("S", 99, 26)]));

    const out = mergeDynamicDeps(graph, dynamic);
    expect(out.changed).toBe(true);
    expect(out.graph).not.toBe(graph);
    expect(out.graph.dependsOn.get(key("S", 1, 2))?.has(key("S", 99, 26))).toBe(true);
    expect(out.graph.dependedBy.get(key("S", 99, 26))?.has(key("S", 1, 2))).toBe(true);
  });

  it("re-detects cycles after dynamic edges are added", () => {
    // Static: A1 and B1 independent. Dynamic adds A1 → B1 and B1 → A1.
    const compiled = buildCompiled([
      { sheet: "S", row: 1, col: 1 },
      { sheet: "S", row: 1, col: 2 }
    ]);
    const graph = buildDependencyGraphFromDeps(compiled);
    expect(graph.circularKeys.size).toBe(0);

    const dynamic = new Map<string, Set<string>>();
    dynamic.set(key("S", 1, 1), new Set([key("S", 1, 2)]));
    dynamic.set(key("S", 1, 2), new Set([key("S", 1, 1)]));

    const out = mergeDynamicDeps(graph, dynamic);
    expect(out.changed).toBe(true);
    expect(out.graph.circularKeys.has(key("S", 1, 1))).toBe(true);
    expect(out.graph.circularKeys.has(key("S", 1, 2))).toBe(true);
  });
});
