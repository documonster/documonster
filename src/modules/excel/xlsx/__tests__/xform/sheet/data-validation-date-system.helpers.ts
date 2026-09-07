/**
 * Shared fixtures for the two halves of the date-validation epoch suite.
 *
 * The buffered assertions are platform-free and run in the browser too; the streaming ones write a real file and
 * live in the `.node.test.ts` sibling. Both make the same claim about the same date, so the date and its two
 * serials are stated once here rather than twice.
 */

/** `2020-01-15`, whose two serials — 43845 and 42383 — are both unremarkable integers. */
export const WHEN = new Date(Date.UTC(2020, 0, 15));

/** Days from 1899-12-30 to {@link WHEN}. */
export const SERIAL_1900 = 43_845;

/** Days from 1904-01-01 to {@link WHEN} — the same day, 1,462 lower. */
export const SERIAL_1904 = SERIAL_1900 - 1462;

/** The bound as a number, however the container gave it back. */
export function boundSerial(value: string | number | Date | undefined, date1904: boolean): number {
  if (typeof value === "number") {
    return value;
  }
  if (typeof value === "string") {
    return Number(value);
  }
  // XLSX reconstructs a `Date`; convert it back the way the writer would.
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return Math.round(((value as Date).getTime() - epoch) / 86_400_000);
}
