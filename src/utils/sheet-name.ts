/**
 * When a sheet name in a formula reference has to be quoted.
 *
 * Excel allows a sheet name unquoted only when it is entirely ASCII letters, digits and
 * underscores and does not start with a digit — any space, dot, comma, bracket, `&`, CJK
 * character or other symbol needs quotes. A name that reads as something else is quoted
 * too: `A1!B2`, `R1C1!B2` and `TRUE!B2` are not references to a sheet, which is why Excel
 * writes `'A1'!B2`. Quoting a name that did not need it is harmless, so the cell test does
 * not bound the column at XFD.
 *
 * Shared by the formula printer (`Formula.format`) and Excel's reference writers, which
 * used to disagree: the printer wrote `A1!B2` and let `.` and `\` through unquoted.
 */
export function sheetNameNeedsQuotes(sheetName: string): boolean {
  return (
    !/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ||
    /^[A-Za-z]{1,3}\d+$/.test(sheetName) ||
    /^(?:[Rr]\d*[Cc]\d*|[RrCc]\d*)$/.test(sheetName) ||
    /^(?:true|false)$/i.test(sheetName)
  );
}

/**
 * `sheetName` as it is written in a formula reference: quoted when
 * {@link sheetNameNeedsQuotes} says so, with an apostrophe inside it doubled.
 *
 * @example quoteSheetName("Sheet1") // "Sheet1"
 * @example quoteSheetName("My Sheet") // "'My Sheet'"
 * @example quoteSheetName("It's Fine") // "'It''s Fine'"
 * @example quoteSheetName("2023 Data") // "'2023 Data'"
 * @example quoteSheetName("A1") // "'A1'"
 */
export function quoteSheetName(sheetName: string): string {
  return sheetNameNeedsQuotes(sheetName) ? `'${sheetName.replace(/'/g, "''")}'` : sheetName;
}
