import type { NumFmt } from "@excel/types";

/**
 * A number format as its code. A style's `numFmt` is the code, or the `{ id, formatCode }` pair `styles.xml` holds;
 * every reader of one wants the code, and an absent one stays absent.
 */
export function numFmtCode(numFmt: string | NumFmt): string;
export function numFmtCode(numFmt: string | NumFmt | undefined): string | undefined;
export function numFmtCode(numFmt: string | NumFmt | undefined): string | undefined {
  return numFmt === undefined || typeof numFmt === "string" ? numFmt : numFmt.formatCode;
}
