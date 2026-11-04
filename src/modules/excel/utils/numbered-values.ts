/** How a values array is indexed: `0` from index `0`, `1` with the index being the row or column number. */
export type ValuesBase = 0 | 1;

/**
 * Options for the strict forms of the values readers and writers — `Row.getValues` / `setValues`,
 * `Column.getValues` / `setValues`, `Worksheet.getValues` and `Stream.getRowValues`.
 */
export interface ValuesOptions {
  /**
   * `0`: plain JavaScript indexing — index `0` is column A (along a row) or row 1 (along a column, or the sheet's
   * rows). `1`: the index *is* the column or row number — index `0` is unused — so it can be handed straight to the
   * `Cell` namespace.
   *
   * A read and a write with the same `base` are inverses.
   */
  base: ValuesBase;
}

/**
 * `values` re-indexed so that the index is the row or column number, index `0` a hole. Shared by
 * `Row.setValues` and `Column.setValues`, whose `{ base }` option means the same thing along either axis.
 */
export function numberedValues<T>(values: readonly T[], base: ValuesBase): T[] {
  const out: T[] = [];
  values.forEach((value, index) => {
    if (index + 1 - base > 0) {
      out[index + 1 - base] = value;
    }
  });
  return out;
}
