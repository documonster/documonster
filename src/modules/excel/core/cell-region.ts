/**
 * A set of cells, per sheet, held as disjoint rectangles rather than as one record per cell.
 *
 * This is what a reference-defined name covers. It used to be a `CellMatrix` — an object per cell — so the size of a
 * name was the size of its area: `Sheet1!$A$1:$A$1048576` built a million objects, a whole-sheet name tried to build
 * seventeen billion and aborted the process, and a whole-column name (`$A:$A`) could not be represented at all and was
 * dropped. Here a name costs one rectangle per area, whatever the area's size.
 *
 * Coordinates are one-based and inclusive, as in `colCache`.
 */
import type { LineEdit } from "@excel/utils/reference-rewrite";
import { editSpan } from "@excel/utils/reference-rewrite";
import { sheetNameKey } from "@utils/name-key";

/** One rectangle of cells. */
export interface CellRect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** Disjoint rectangles per sheet, in the order the sheets were first named. */
export type CellRegion = Map<string | undefined, CellRect[]>;

export function createCellRegion(): CellRegion {
  return new Map();
}

function overlaps(a: CellRect, b: CellRect): boolean {
  return a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
}

/** `a` minus `b`, as up to four disjoint rectangles. */
function subtract(a: CellRect, b: CellRect): CellRect[] {
  if (!overlaps(a, b)) {
    return [a];
  }
  const pieces: CellRect[] = [];
  if (a.top < b.top) {
    pieces.push({ top: a.top, left: a.left, bottom: b.top - 1, right: a.right });
  }
  if (b.bottom < a.bottom) {
    pieces.push({ top: b.bottom + 1, left: a.left, bottom: a.bottom, right: a.right });
  }
  const top = Math.max(a.top, b.top);
  const bottom = Math.min(a.bottom, b.bottom);
  if (a.left < b.left) {
    pieces.push({ top, left: a.left, bottom, right: b.left - 1 });
  }
  if (b.right < a.right) {
    pieces.push({ top, left: b.right + 1, bottom, right: a.right });
  }
  return pieces;
}

/** Join `rect` with a neighbour it extends exactly, repeatedly, so cell-by-cell additions stay a few rectangles. */
function insertMerged(rects: CellRect[], rect: CellRect): void {
  let current = rect;
  for (let merged = true; merged;) {
    merged = false;
    for (let i = 0; i < rects.length; i++) {
      const other = rects[i];
      const sameColumns = other.left === current.left && other.right === current.right;
      const sameRows = other.top === current.top && other.bottom === current.bottom;
      if (
        (sameColumns && (other.bottom + 1 === current.top || current.bottom + 1 === other.top)) ||
        (sameRows && (other.right + 1 === current.left || current.right + 1 === other.left))
      ) {
        current = {
          top: Math.min(other.top, current.top),
          left: Math.min(other.left, current.left),
          bottom: Math.max(other.bottom, current.bottom),
          right: Math.max(other.right, current.right)
        };
        rects.splice(i, 1);
        merged = true;
        break;
      }
    }
  }
  rects.push(current);
}

/** Add the cells of `rect` on `sheet`. */
export function cellRegionAdd(region: CellRegion, sheet: string | undefined, rect: CellRect): void {
  let rects = region.get(sheet);
  if (!rects) {
    rects = [];
    region.set(sheet, rects);
  }
  let pieces = [{ ...rect }];
  for (const existing of rects) {
    pieces = pieces.flatMap(piece => subtract(piece, existing));
  }
  for (const piece of pieces) {
    insertMerged(rects, piece);
  }
}

/** Remove the cells of `rect` on `sheet`. `true` when any of them was in the region. */
export function cellRegionRemove(
  region: CellRegion,
  sheet: string | undefined,
  rect: CellRect
): boolean {
  const rects = region.get(sheet);
  if (!rects?.some(existing => overlaps(existing, rect))) {
    return false;
  }
  const kept = rects.flatMap(existing => subtract(existing, rect));
  rects.length = 0;
  for (const piece of kept) {
    insertMerged(rects, piece);
  }
  return true;
}

export function cellRegionContains(
  region: CellRegion,
  sheet: string | undefined,
  row: number,
  col: number
): boolean {
  return (
    region
      .get(sheet)
      ?.some(r => r.top <= row && row <= r.bottom && r.left <= col && col <= r.right) ?? false
  );
}

/**
 * Carry the region's rectangles on `edit.sheet` through a row or column edit, by the rules every other reference
 * follows (see `reference-rewrite`): a rectangle an insertion falls strictly inside grows, one the deletion covers goes,
 * one it overlaps shrinks. Sheets are matched as Excel matches their names.
 *
 * It used to treat the region as a set of cells, so an insertion inside a name's area cut it in two around the new
 * lines — after `Row.insert` a name over `$A$1:$A$4` covered `$A$1:$A$2,$A$4:$A$5`, where Excel has `$A$1:$A$5`.
 */
export function cellRegionEdit(region: CellRegion, edit: LineEdit): void {
  const [low, high] =
    edit.kind === "rows" ? (["top", "bottom"] as const) : (["left", "right"] as const);
  const key = sheetNameKey(edit.sheet);
  for (const [sheet, rects] of region) {
    if (sheet === undefined || sheetNameKey(sheet) !== key) {
      continue;
    }
    const moved: CellRect[] = [];
    for (const rect of rects) {
      const span = editSpan(rect[low], rect[high], edit);
      if (span !== undefined) {
        moved.push({ ...rect, [low]: span[0], [high]: span[1] });
      }
    }
    rects.length = 0;
    for (const rect of moved) {
      insertMerged(rects, rect);
    }
  }
}

/** Whether the region covers any cell. */
export function cellRegionIsEmpty(region: CellRegion): boolean {
  for (const rects of region.values()) {
    if (rects.length > 0) {
      return false;
    }
  }
  return true;
}

/**
 * The same region with sheet `from` renamed `to` — or, with `to` undefined, without that sheet. A new map, in the same
 * sheet order.
 */
export function cellRegionRenameSheet(
  region: CellRegion,
  from: string,
  to: string | undefined
): CellRegion {
  const key = sheetNameKey(from);
  const next: CellRegion = new Map();
  for (const [sheet, rects] of region) {
    if (sheet !== undefined && sheetNameKey(sheet) === key) {
      if (to !== undefined) {
        next.set(to, rects);
      }
    } else {
      next.set(sheet, rects);
    }
  }
  return next;
}

/**
 * The region as rectangles in a canonical order: sheets as first named, then each rectangle by its top-left cell in
 * row-major order. Each rectangle is grown down from that cell first, then across — so a set of cells has one
 * decomposition however it was built, and it is the one a name's ranges have always been written as.
 *
 * Computed on the grid the rectangles' own edges cut the sheet into, not cell by cell: every block of that grid is
 * wholly inside the region or wholly outside it, so growing over blocks gives the same rectangles as growing over
 * cells, at a cost set by the number of rectangles rather than their area.
 */
export function cellRegionRects(
  region: CellRegion
): { sheet: string | undefined; rect: CellRect }[] {
  const result: { sheet: string | undefined; rect: CellRect }[] = [];
  for (const [sheet, rects] of region) {
    if (rects.length === 0) {
      continue;
    }
    const rowEdges = edges(rects, "top", "bottom");
    const colEdges = edges(rects, "left", "right");
    const covered = new Set<number>();
    const width = colEdges.length;
    for (const rect of rects) {
      const r0 = rowEdges.indexOf(rect.top);
      const r1 = rowEdges.indexOf(rect.bottom + 1);
      const c0 = colEdges.indexOf(rect.left);
      const c1 = colEdges.indexOf(rect.right + 1);
      for (let r = r0; r < r1; r++) {
        for (let c = c0; c < c1; c++) {
          covered.add(r * width + c);
        }
      }
    }
    const order = [...covered].sort((a, b) => a - b);
    const has = (r: number, c: number): boolean => covered.has(r * width + c);
    for (const key of order) {
      if (!covered.has(key)) {
        continue;
      }
      const r0 = Math.floor(key / width);
      const c0 = key % width;
      let r1 = r0;
      while (has(r1 + 1, c0)) {
        r1++;
      }
      let c1 = c0;
      const columnFull = (c: number): boolean => {
        for (let r = r0; r <= r1; r++) {
          if (!has(r, c)) {
            return false;
          }
        }
        return true;
      };
      while (columnFull(c1 + 1)) {
        c1++;
      }
      for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
          covered.delete(r * width + c);
        }
      }
      result.push({
        sheet,
        rect: {
          top: rowEdges[r0],
          left: colEdges[c0],
          bottom: rowEdges[r1 + 1] - 1,
          right: colEdges[c1 + 1] - 1
        }
      });
    }
  }
  return result;
}

function edges(
  rects: readonly CellRect[],
  low: "top" | "left",
  high: "bottom" | "right"
): number[] {
  const values = new Set<number>();
  for (const rect of rects) {
    values.add(rect[low]);
    values.add(rect[high] + 1);
  }
  return [...values].sort((a, b) => a - b);
}
