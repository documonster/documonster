/**
 * A built-in number format a read can name only by its id.
 *
 * A style's `numFmt` is a format code, and most built-in ids have one (`14` is `mm-dd-yy`). The rest do not: 5–8
 * are the currency formats, 41–44 the accounting ones, 27–36 and 50–81 the East Asian and Thai dates, times and
 * numbers, 23–26 reserved — each displayed in the form of the reader's locale (`display-locale.ts`), so there is no
 * one code to give them, and guessing one would state a format the file never did. Those ids used to be dropped on
 * read, so the cell read as General and was written back as General: a currency column lost its currency on a round
 * trip, with nothing reported.
 *
 * Such an id is set aside on the style instead (`@utils/unrecognized-values`): `numFmt` reads as absent, the id is
 * written back while it is, and the read report lists it (`unrecognizedValues`, path `….style.numFmt`). Setting
 * `numFmt` takes over, as for any kept value. An id that shows a date or a time in every locale defining it (27–36,
 * 50–58, 71–81) still makes its cell's number a date, as `mm-dd-yy` does: the form is the locale's, that it is a date
 * is not.
 *
 * A custom id (164 and up) whose `<numFmt>` is missing is not kept. It names no format in the file it came from, and
 * a writer numbers its custom formats afresh, so written back it would name whichever format took that number. It
 * reads as General, and the read reports it in `readWithDiagnostics().lost`.
 *
 * Four states of a number format, then, and every reader and writer holds them the same way: **inherited** (`numFmt`
 * absent, nothing kept — a sheet's format takes the workbook default's, `Workbook.getDefaultStyle`), an **explicit
 * code** (`numFmt` set; General is a value, not an absence, and is written as one), a **kept built-in id** (above),
 * and a **dangling reference** (read as General, reported, never written).
 */
import { defaultNumFormats } from "@excel/xlsx/defaultnumformats";
import { isDateFmt } from "@utils/number-format";
import { setAsideValue, unrecognizedValue } from "@utils/unrecognized-values";

/** The first number-format id a workbook defines itself; the ones below are reserved for the built-ins. */
export const FIRST_CUSTOM_NUM_FMT_ID = 164;

/** Whether `id` is a reserved built-in number-format id other than General. */
export function isBuiltInNumFmtId(id: number): boolean {
  return Number.isInteger(id) && id > 0 && id < FIRST_CUSTOM_NUM_FMT_ID;
}

/** Keep the built-in number-format `id` on `style`, whose `numFmt` the reader leaves absent. */
export function keepNumFmtId(style: object, id: number): void {
  setAsideValue(style, "numFmt", String(id));
}

/** The built-in number-format id kept on `style`, while its `numFmt` is still absent. */
export function keptNumFmtId(style: object | undefined): number | undefined {
  const text = unrecognizedValue(style, "numFmt");
  if (text === undefined) {
    return undefined;
  }
  const id = Number(text);
  return isBuiltInNumFmtId(id) ? id : undefined;
}

const dateIds = new Map<number, boolean>();

/**
 * Whether a number format shows a number as a date: by its code, or — a format with none — by the built-in id kept for
 * want of one. What every reader asks to decide whether a cell's serial number is a date.
 */
export function showsDate(code: string | undefined, keptId: number | undefined): boolean {
  return code !== undefined ? isDateFmt(code) : keptId !== undefined && isDateNumFmtId(keptId);
}

/**
 * Whether the built-in number-format `id` shows a date or a time in every locale that defines it (27–36, 50–58 and
 * 71–81 do; the currencies 5–8 do not). A cell formatted with one holds a date as a serial number, as one formatted
 * `mm-dd-yy` does, and is read as one: which form the date takes is the locale's, that it is one is not.
 */
export function isDateNumFmtId(id: number): boolean {
  let isDate = dateIds.get(id);
  if (isDate === undefined) {
    const codes = Object.values(defaultNumFormats[id] ?? {});
    isDate = codes.length > 0 && codes.every(code => isDateFmt(code));
    dateIds.set(id, isDate);
  }
  return isDate;
}
