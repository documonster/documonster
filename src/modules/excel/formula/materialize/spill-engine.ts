/**
 * Spill Engine — Persistent state management for dynamic array formulas.
 *
 * Manages persistent spill metadata that survives across
 * `calculateFormulas()` invocations.
 */

import type { SpillRegion, WorkbookLike } from "./types";

// ============================================================================
// Persistent State (WeakMap keyed by workbook — survives across invocations)
// ============================================================================

/**
 * Persistent spill metadata: survives across calculateFormulas invocations.
 * Key: "ws:<id>!row:col" of the source cell → SpillRegion.
 * Stored in a WeakMap keyed by the workbook object to allow GC.
 */
const persistentSpillRegions = new WeakMap<WeakKey, Map<string, SpillRegion>>();

export function getPersistentSpillMap(workbook: WorkbookLike): Map<string, SpillRegion> {
  let map = persistentSpillRegions.get(workbook);
  if (!map) {
    map = new Map();
    persistentSpillRegions.set(workbook, map);
  }
  return map;
}

/**
 * Persistent snapshot of values written to ghost (spill target) cells.
 * Key: ghost cell key "ws:<id>!row:col" → raw value written by spill.
 * Used to detect whether a ghost cell has been modified by the user since
 * the last calculation.
 */
const persistentGhostSnapshots = new WeakMap<WeakKey, Map<string, unknown>>();

export function getGhostSnapshots(workbook: WorkbookLike): Map<string, unknown> {
  let map = persistentGhostSnapshots.get(workbook);
  if (!map) {
    map = new Map();
    persistentGhostSnapshots.set(workbook, map);
  }
  return map;
}
