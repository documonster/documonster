import type { Style } from "@excel/types";
import { copyStyleFacet } from "@excel/utils/copy-style";

/**
 * The single copy of a row's or column's style facets that every cell it covers
 * points at.
 *
 * Without it each cell kept its own deep copy of the same facets — measured at 382
 * bytes per cell, so a 200k-cell sheet carrying one logical style paid 73 MB to say
 * the same thing 200,000 times.
 *
 * **A copy, not the owner's own facet object.** That is what leaves the owner's
 * semantics untouched: `Column.getStyle(ws, 1).font.bold = false` must not reach
 * cells that were already styled, exactly as it does not today. Cells point into the
 * snapshot; a caller mutating the column mutates the column's own object.
 *
 * **Keyed by the style object, not the owner.** Ten places replace a row's or
 * column's `style` wholesale (`applyDefn`, `rowSetModel`, `duplicateRow`,
 * `spliceRows`, …). Keying on the object means every one of those invalidates by
 * construction — a field on the owner would have needed all ten to remember, and the
 * one that forgot would hand cells a snapshot of a style the owner no longer has.
 * Only mutating a facet *on the same object* needs {@link invalidateSharedCellStyle},
 * and those sites are the handful of `set*` calls in this module's callers.
 */
const snapshots = new WeakMap<Partial<Style>, Partial<Style>>();

/**
 * Freeze a snapshot facet and everything under it.
 *
 * A shared facet is reachable from hundreds of thousands of cells, so mutating one
 * in place would silently rewrite all of them. Freezing turns that into a `TypeError`
 * at the offending line — ESM is strict mode, so the assignment throws rather than
 * being ignored. It is the enforcement half of the sharing invariant: `cellOwnStyle`
 * hands out private copies, and this guarantees anything that skipped it fails loudly
 * instead of corrupting a neighbour.
 *
 * `Cell.view` is the case that needs it most: it is documented as a read-only
 * projection and its `font`/`alignment` are the shared objects, deliberately not
 * materialised so that iterating a row stays allocation-free.
 *
 * Shape-agnostic on purpose — the per-facet nesting is already spelled out once in
 * `copy-style.ts` and repeating it here would be a second definition to drift. Runs
 * once per snapshot, not per cell. The `isFrozen` check also terminates on a cycle.
 */
function deepFreeze(value: unknown): void {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const key of Object.keys(value)) {
      deepFreeze((value as Record<string, unknown>)[key]);
    }
  }
}

/**
 * The facet a cell covered by `style`'s owner should point at: the same object for
 * every such cell, built on first use and frozen.
 *
 * `numFmt` and `styleName` are primitives, so they are returned as-is and never enter
 * a snapshot.
 */
export function sharedCellFacet<K extends keyof Style>(
  style: Partial<Style>,
  key: K
): Style[K] | undefined {
  const value = style[key];
  if (!value || typeof value !== "object") {
    return value;
  }
  let snapshot = snapshots.get(style);
  if (!snapshot) {
    snapshot = {};
    snapshots.set(style, snapshot);
  }
  let facet = snapshot[key];
  if (facet === undefined) {
    facet = copyStyleFacet(key, value);
    deepFreeze(facet);
    snapshot[key] = facet;
  }
  return facet as Style[K];
}

/**
 * The frozen style container every cell resolving to `style` should point at.
 *
 * A cell whose style comes wholly from one source — every cell that loaded with a
 * given `styleId`, since the reader caches one style model per id — needs no container
 * of its own. Sharing it is worth ~85 bytes per cell on top of the facets, and it is
 * the last per-cell cost of a style that many cells agree on.
 *
 * Frozen for the same reason the facets are: it both enforces the sharing and marks
 * it, so `ownStyleContainer` can recognise a shared container without a second flag
 * and a write that skips it throws rather than rewriting every cell that shares it.
 *
 * The container's own facets are the snapshots, so `sharedCellFacet` and this agree by
 * construction. Fields that are not facets — `numFmt`, `styleName`, and the xf-level
 * `pivotButton`/`apply*` flags the reader preserves — are carried across verbatim.
 */
export function sharedCellContainer(style: Partial<Style>): Partial<Style> {
  let container = containers.get(style);
  if (!container) {
    container = { ...style };
    for (const key of OBJECT_FACETS) {
      if (container[key] !== undefined) {
        container[key] = sharedCellFacet(style, key) as never;
      }
    }
    Object.freeze(container);
    containers.set(style, container);
  }
  return container;
}

const containers = new WeakMap<Partial<Style>, Partial<Style>>();

/**
 * The container an unstyled cell points at.
 *
 * A cell always has a `style` object because 182 readers say `cell.style.font`, but a
 * cell with no style has nothing to put in it, and allocating one per cell is pure
 * overhead — 200k of them on a plain sheet. Frozen like every shared container, so
 * `ownStyleContainer` gives a cell its own the moment anything writes to it.
 */
export const EMPTY_CELL_STYLE: Partial<Style> = Object.freeze({});

/** The style facets that are objects, and so the ones that can be shared or copied. */
const OBJECT_FACETS = ["font", "alignment", "protection", "border", "fill"] as const;

/**
 * Drop the snapshot of `style`, so the next cell that needs a facet takes a fresh one.
 *
 * Needed when a facet is replaced **on this same object** — a wholesale
 * `owner.style = …` invalidates on its own, because the snapshot is keyed by object.
 * Call it in two situations:
 *
 *  - A `set*` that writes `owner.style.font = …` in place.
 *  - Handing the style to a caller, who may mutate it in place. Today a cell created
 *    after such a mutation copies the mutated facet, and invalidating on handout is
 *    what preserves that.
 */
export function invalidateSharedCellStyle(style: Partial<Style>): void {
  snapshots.delete(style);
  containers.delete(style);
}
