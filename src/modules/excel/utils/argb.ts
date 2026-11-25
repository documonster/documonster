/**
 * A colour's ARGB as a file writes it. `rgb` is `ST_UnsignedIntHex` — four bytes — and a `BrtColor` holds alpha too, so
 * a six-digit colour, which states none, is the opaque one: the only sensible reading, and the one every writer makes.
 * One function for all of them: the XLSX style writer, the XLSB one and the sparkline writer each padded on their own,
 * and the sparkline writer did not, writing `rgb="C00000"` the schema refuses.
 */
export function opaqueArgb(argb: string): string {
  return argb.length === 6 ? `FF${argb}` : argb;
}
