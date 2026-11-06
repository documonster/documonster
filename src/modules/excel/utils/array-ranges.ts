import { decodeCell } from "@excel/utils/address";

/**
 * Array formula ranges a sheet writer has claimed so far, and the one rule every writer applies to them.
 *
 * A cell is one record, so it can belong to one array. Two array ranges that intersect — a later master inside an
 * earlier range, a later range crossing or covering an earlier one — would make some cell a member of both, and Excel
 * refuses the file. Every writer therefore keeps the earlier range whole and writes the later array as its master's
 * cached result alone, with no followers, and reports the formula as lost. Nothing is reassigned silently.
 */

/** An array formula's range as written: zero-based bounds, the master's address and the `ref` it was given. */
export interface ArrayRange {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
  readonly address: string;
  readonly ref: string;
}

/** `"E1:E4"` as zero-based bounds, or `undefined` when it is not a range a writer can read. */
export function parseArrayRefBounds(
  ref: string
): { top: number; left: number; bottom: number; right: number } | undefined {
  const [start, end = start] = ref.split(":");
  const first = tryDecode(start);
  const last = tryDecode(end);
  if (first === undefined || last === undefined) {
    return undefined;
  }
  return {
    top: Math.min(first.row, last.row),
    left: Math.min(first.column, last.column),
    bottom: Math.max(first.row, last.row),
    right: Math.max(first.column, last.column)
  };
}

/** The earlier array range containing the zero-based cell, if any. */
export function enclosingArrayRange(
  ranges: readonly ArrayRange[],
  row: number,
  column: number
): ArrayRange | undefined {
  return ranges.find(
    range =>
      row >= range.top && row <= range.bottom && column >= range.left && column <= range.right
  );
}

/**
 * Claim the range of the array formula mastered at `address`, in row-major order.
 *
 * Returns the earlier range it intersects — its master cell or any cell of `ref` — in which case nothing is recorded
 * and the formula cannot be written; or `undefined` once the range has been recorded. A `ref` that is not a range is
 * not recorded, and only its master cell is tested.
 */
export function claimArrayRange(
  ranges: ArrayRange[],
  address: string,
  ref: string
): ArrayRange | undefined {
  const master = tryDecode(address);
  if (master === undefined) {
    return undefined;
  }
  const bounds = parseArrayRefBounds(ref);
  const outer =
    enclosingArrayRange(ranges, master.row, master.column) ??
    (bounds === undefined
      ? undefined
      : ranges.find(
          range =>
            bounds.top <= range.bottom &&
            bounds.bottom >= range.top &&
            bounds.left <= range.right &&
            bounds.right >= range.left
        ));
  if (outer !== undefined) {
    return outer;
  }
  if (bounds !== undefined) {
    ranges.push({ ...bounds, address, ref });
  }
  return undefined;
}

/** How an array formula lost to an enclosing range is reported, in either container. */
export function overlappingArrayReason(outer: ArrayRange): string {
  return `array formula (overlaps the array range ${outer.ref} of ${outer.address})`;
}

function tryDecode(address: string): { row: number; column: number } | undefined {
  try {
    const decoded = decodeCell(address);
    return Number.isFinite(decoded.r) && Number.isFinite(decoded.c)
      ? { row: decoded.r, column: decoded.c }
      : undefined;
  } catch {
    return undefined;
  }
}
