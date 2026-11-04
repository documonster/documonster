/**
 * Arguments Excel reads from their first cell, stretched to the size of another argument. The
 * evaluator stretches the reference at run time and dependency extraction stretches the bound
 * one, so both agree on which cells are read. What Excel does is measured, not documented beyond
 * the first rule (`src/modules/formula/__tests__/reference/excel-ground-truth.test.ts`):
 *
 * - `SUMIF` / `AVERAGEIF` `sum_range`: the criteria range's size and shape (Microsoft: "a range
 *   of cells that starts with the first cell in sum_range but has the same dimensions as range").
 * - `LOOKUP` `result_vector`: as many cells as `lookup_vector`, down a column when it is one
 *   column of two or more cells, otherwise across a row — a single cell runs across whichever
 *   way the lookup runs. A two-dimensional `result_vector` is #N/A.
 *
 * Either is cut at the sheet's last row and column; Excel reads the cells that exist.
 */

/** The argument each function stretches. */
export const STRETCHED_ARGUMENT: ReadonlyMap<string, number> = new Map([
  ["SUMIF", 2],
  ["AVERAGEIF", 2],
  ["LOOKUP", 2]
]);

export const LAST_ROW = 1_048_576;
export const LAST_COL = 16_384;

export interface StretchedArea {
  readonly sheet: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** `height` × `width` cells from (`top`, `left`) on `sheet`, cut at the sheet's last row and column. */
export function stretchArea(
  sheet: string,
  top: number,
  left: number,
  height: number,
  width: number
): StretchedArea {
  return {
    sheet,
    top,
    left,
    bottom: Math.min(top + height - 1, LAST_ROW),
    right: Math.min(left + width - 1, LAST_COL)
  };
}

/**
 * The shape `LOOKUP` reads a `result_vector` of `height` × `width` in, for a lookup of `length`
 * cells; `undefined` for a two-dimensional one.
 */
export function lookupResultShape(
  height: number,
  width: number,
  length: number
): { height: number; width: number } | undefined {
  if (height > 1 && width > 1) {
    return undefined;
  }
  return height > 1 ? { height: length, width: 1 } : { height: 1, width: length };
}
