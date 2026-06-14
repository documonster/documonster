/**
 * `Range` namespace surface — geometric range helpers.
 *
 * `import { Range } from "documonster/excel"` → `Range.create("A1:B2")`,
 * `Range.contains(r, "A1")`, `Range.forEachAddress(r, cb)`.
 */
export {
  rangeCreate as create,
  rangeContains as contains,
  rangeContainsEx as containsCell,
  rangeIntersects as intersects,
  rangeForEachAddress as forEachAddress,
  rangeExpand as expand,
  rangeExpandToAddress as expandToAddress,
  rangeToString as toString,
  rangeCount as count
} from "@excel/range";

/** A range handle. */
export type { RangeData as Handle } from "@excel/range";
