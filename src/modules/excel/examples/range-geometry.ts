/**
 * Range geometry helpers — the `Range` namespace.
 *
 * Covers the geometric range utilities (no worksheet required):
 * - Range.create        — build a range from an A1 range string
 * - Range.toString      — serialise back to a range string
 * - Range.count         — number of cells covered
 * - Range.contains      — does the range contain an A1 address
 * - Range.containsCell  — does the range contain a decoded Address
 * - Range.intersects    — do two ranges overlap
 * - Range.expand        — grow a range to include a t/l/b/r box
 * - Range.expandToAddress — grow a range to include an A1 address
 * - Range.forEachAddress — iterate every cell address in the range
 *
 * Usage:
 *   npx tsx src/modules/excel/examples/range-geometry.ts
 */
import { Range } from "@excel/index";
import type { Address } from "@excel/types";

// 1. Create + serialise + count
const r = Range.create("B2:D5");
console.log("range:", Range.toString(r)); // B2:D5
console.log("cell count:", Range.count(r)); // 3 cols * 4 rows = 12

// 2. contains (by A1 string)
console.log("contains C3:", Range.contains(r, "C3")); // true
console.log("contains A1:", Range.contains(r, "A1")); // false

// 3. containsCell (by an Address object: 1-based col/row)
const cell: Address = { address: "D5", col: 4, row: 5 };
console.log("containsCell D5:", Range.containsCell(r, cell)); // true

// 4. intersects (two ranges)
const other = Range.create("C4:F8");
const disjoint = Range.create("H1:I2");
console.log("intersects C4:F8:", Range.intersects(r, other)); // true
console.log("intersects H1:I2:", Range.intersects(r, disjoint)); // false

// 5. expand — grow the box to include row 7 / col 6 (F)
Range.expand(r, 2, 2, 7, 6);
console.log("after expand:", Range.toString(r)); // B2:F7

// 6. expandToAddress — grow to include H10
Range.expandToAddress(r, "H10");
console.log("after expandToAddress H10:", Range.toString(r)); // B2:H10

// 7. forEachAddress — enumerate cells of a small range
const small = Range.create("A1:B2");
const addresses: string[] = [];
Range.forEachAddress(small, (address, row, col) => {
  addresses.push(`${address}(r${row},c${col})`);
});
console.log("forEachAddress A1:B2:", addresses.join(" "));

console.log("Done.");
