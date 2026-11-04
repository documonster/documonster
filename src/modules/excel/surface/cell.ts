/**
 * `Cell` namespace surface — public `(ws, addr, …)` cell operations.
 *
 * Consumed as `import { Cell } from "documonster/excel"` → `Cell.setValue(ws, "A1", 42)`.
 * Each function resolves the cell handle via `getCell(ws, addr)` and delegates
 * to the internal handle-level helpers. Consumers never hold a `CellData`.
 *
 * The `addr` parameter is either an `"A1"` string or a 1-based row number
 * followed by a column (`Cell.setValue(ws, 1, 1, 42)`), hence the
 * `string | number` unions in the signatures below.
 *
 * This is a flat-named-export module; `excel/index.ts` re-exports it via
 * `export * as Cell`, which tree-shakes per-member on rolldown / rspack.
 */
import type {
  CellData,
  CellModel,
  CellView,
  CellValueInputType,
  CellValueType,
  FormulaResult,
  NoteConfig
} from "@excel/core/cell";
import {
  checkCellValue,
  checkFormulaResult,
  checkSharedFormulaAt,
  cellAddName,
  cellAlignment,
  cellDateKind,
  cellGetDateParts,
  cellGetTemporal,
  cellSetDateParts,
  validateDateParts,
  cellBorder,
  cellComment,
  cellDisplayText,
  cellEffectiveType,
  cellFill,
  cellFont,
  cellFormula,
  cellGetModel,
  cellGetStyle,
  cellOwnStyle,
  cellGetValue,
  cellHyperlink,
  cellIsMerged,
  cellMaster,
  cellNote,
  cellNumFmt,
  cellProtection,
  cellResult,
  cellSetAlignment,
  cellSetBorder,
  cellSetComment,
  cellSetDataValidation,
  cellSetFill,
  cellSetFont,
  cellSetModel,
  cellSetName,
  cellSetNames,
  cellSetNote,
  cellSetNumFmt,
  cellSetProtection,
  cellSetResult,
  cellSetStyle,
  cellSetValue,
  cellText,
  cellType,
  cellView
} from "@excel/core/cell";
import { dataValidationFind } from "@excel/core/data-validations";
import {
  definedNamesCheckName,
  definedNamesGetNamesEx,
  definedNamesRemoveAllNames,
  definedNamesRemoveEx
} from "@excel/core/defined-names";
import { ValueType } from "@excel/core/enums";
import { isNoteData } from "@excel/core/note";
import type { NoteData } from "@excel/core/note";
import type { TemporalKind, TemporalPlainValue } from "@excel/core/temporal";
import { getCellStyle } from "@excel/core/workbook-core";
import {
  checkWorksheet,
  checkWorksheetAttached,
  getCell,
  getSheetWorkbook,
  inheritedNumFmt,
  locateCell,
  previewCell,
  resolveAddress
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { ExcelArgumentError, InvalidAddressError } from "@excel/errors";
import { checkDataValidationRule, checkNote } from "@excel/surface/argument-checks";
import { checkCellModel } from "@excel/surface/model-checks";
import type {
  Alignment,
  Borders,
  DecodedAddress,
  DataValidationRule,
  Fill,
  Font,
  NumFmt,
  Protection,
  Style
} from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import type { DateFormatKind } from "@excel/utils/cell-format";
import { copyStyle } from "@excel/utils/copy-style";
import { checkStyle, checkStyleFacet } from "@excel/utils/style-checks";
import type { ExcelDateTimeParts } from "@utils/excel-serial";

/** @deprecated Use `Worksheet.Handle`. Removed in 1.0. */
export type Sheet = WorksheetData;

/**
 * Resolve the `(ws, addr)` / `(ws, row, col)` reader overload pair.
 *
 * `argc` is the caller's `arguments.length`: the 3-arg form addresses by
 * `(row, col)`, the 2-arg form by `"A1"`. Passing `col` positionally instead of
 * inspecting `argc` would make `(ws, "A1", 99)` silently address a completely
 * different cell (`n2l(99) + "A1"`), which is exactly what the old
 * `addr: string | number, col?: number` signature allowed.
 */
function target(
  ws: WorksheetData,
  addr: string | number,
  col: number | undefined,
  argc: number
): CellData {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  return argc >= 3 ? getCell(ws, rowNumber(addr), col) : getCell(ws, addr);
}

/**
 * The row of the `(row, col)` form, which must be a number. Handed `(ws, "A1", 5)`, the resolver built the address
 * `n2l(5) + "A1"` — `EA1`, column 131 — so a stray extra argument silently addressed a cell far from the one named.
 */
function rowNumber(addr: string | number): number {
  if (typeof addr !== "number") {
    throw new InvalidAddressError(
      String(addr),
      'the (row, col) form takes a row number; an "A1" address takes no column'
    );
  }
  return addr;
}

/**
 * {@link target} for a reader that answers with a value: resolves the same overload pair and fails on the
 * same malformed references, but never creates the row or the cell — `undefined` when nothing is there,
 * and the reader answers as for an empty cell.
 *
 * A read must not change what the workbook saves — materialising widened `rowCount` and `<dimension>`,
 * and in a styled row or column wrote a styled empty cell into the file. {@link getStyle} is the one
 * deliberate exception: it hands back the live style record for the caller to edit in place
 * (`Cell.getStyle(ws, "B2").font = …`), and there is nothing to edit until the cell exists.
 */
function peek(
  ws: WorksheetData,
  addr: string | number,
  col: number | undefined,
  argc: number
): CellData | undefined {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  return argc >= 3 ? locateCell(ws, rowNumber(addr), col) : locateCell(ws, addr);
}

/** The overload pair resolved to an address, for a reader that answers from the address alone. */
function addressOf(
  ws: WorksheetData,
  addr: string | number,
  col: number | undefined,
  argc: number
): DecodedAddress {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  return argc >= 3 ? resolveAddress(ws, rowNumber(addr), col) : resolveAddress(ws, addr);
}

/**
 * The cell for a reader that hands a *style facet* back to the caller, without creating it.
 *
 * A cell's facets may be shared — with the row/column snapshot it inherited, with
 * sibling cells sharing that snapshot, or with whoever passed the facet to a setter —
 * so a reader that lets the caller reach one has to give the cell its own copy first
 * or `Cell.getFont(a).bold = x` would reach every cell in the column.
 *
 * This is the only boundary where that happens. Core internals and both writers read
 * the shared facets directly, which is what keeps a styled sheet's memory flat.
 *
 * An existing cell's facet stays live, as it always was. An absent cell is previewed
 * rather than created ({@link previewCell}): the reader reports what the cell would
 * inherit, and editing that copy changes nothing — {@link getStyle} is the reader that
 * creates the cell for editing.
 */
function own(
  ws: WorksheetData,
  addr: string | number,
  col: number | undefined,
  argc: number
): CellData {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  const cell = argc >= 3 ? previewCell(ws, rowNumber(addr), col) : previewCell(ws, addr);
  cellOwnStyle(cell);
  return cell;
}

/**
 * Resolve the `(ws, addr, value)` / `(ws, row, col, value)` writer overload
 * pair. See {@link target} for why `argc` decides.
 */
/**
 * The cell and the value of a setter's `(addr, value)` or `(row, col, value)` arguments. `check` sees the value
 * before the cell is looked up, so a refused value does not leave a new empty cell behind.
 */
function targetWithValue<V>(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: V | number,
  value: V | undefined,
  argc: number,
  check?: (value: V) => void
): [CellData, V] {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  const resolved = (argc >= 4 ? value : valueOrCol) as V;
  check?.(resolved);
  return [
    argc >= 4 ? getCell(ws, rowNumber(addr), valueOrCol as number) : getCell(ws, addr),
    resolved
  ];
}

/**
 * A name a cell is given is a defined name, so it is checked as `DefinedNames.add` checks one — before the cell is
 * looked up, so a refused name leaves nothing behind.
 */
function checkCellName(ws: WorksheetData, where: string, name: unknown): void {
  definedNamesCheckName(getSheetWorkbook(ws)._definedNames, where, name);
}

// --- value / type / text ---

/** Read a cell value by `"A1"` address. */
export function getValue(ws: WorksheetData, addr: string): CellValueType;
/** Read a cell value by 1-based (row, col). */
export function getValue(ws: WorksheetData, row: number, col: number): CellValueType;
export function getValue(ws: WorksheetData, addr: string | number, col?: number): CellValueType {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellGetValue(cell) : null;
}
/**
 * Set a cell value by "A1" address.
 *
 * A value of a known shape is checked as that shape, and a malformed one is refused with `ExcelArgumentError`: rich
 * text (`{ richText: [{ text, font? }] }`), a formula or shared formula (non-empty text; a shared formula's master is a
 * formula cell above or to the left), its `result` (a scalar, a `Date` or `{ error }`), a hyperlink (`hyperlink`,
 * `text` and `tooltip` are text), an error (`{ error: "#N/A" }` — one of Excel's error literals) and a checkbox. An
 * object or array of none of those shapes is stored as its JSON text, as it always has been.
 */
export function setValue(ws: WorksheetData, addr: string, value: CellValueInputType): void;
/** Set a cell value by 1-based (row, col). */
export function setValue(
  ws: WorksheetData,
  row: number,
  col: number,
  value: CellValueInputType
): void;
export function setValue(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: CellValueInputType,
  value?: CellValueInputType
): void {
  const argc = arguments.length;
  const [cell, resolved] = targetWithValue<CellValueInputType>(
    ws,
    addr,
    valueOrCol,
    value,
    argc,
    v => {
      const site = { where: "Cell.setValue", field: "value" };
      checkCellValue(v, site);
      // A shared formula's master is checked before the cell is created, so a refusal leaves nothing behind.
      const at =
        argc >= 4
          ? addressOf(ws, addr, valueOrCol as number, 3)
          : addressOf(ws, addr, undefined, 2);
      checkSharedFormulaAt(ws, at.row, at.col, v, site);
    }
  );
  cellSetValue(cell, resolved);
}
// --- dates as calendar values ---

/**
 * Read a date cell as calendar fields, or `undefined` if it holds no date.
 *
 * The unambiguous counterpart of {@link getValue} for a date. `getValue` returns a `Date`, which is an *instant*
 * and therefore needs a convention to be read as a calendar value — this library's is that the civil fields are
 * the `Date`'s **UTC** fields, and reading it with `getFullYear()` instead of `getUTCFullYear()` gives the wrong
 * day in half the world's timezones. These fields cannot be read wrong: they have no timezone.
 *
 * Available on every supported runtime. `Temporal.PlainDate.from(parts)` converts one in a line, which is why
 * this rather than {@link getTemporal} is the primitive.
 *
 * @example
 * ```ts
 * import { Cell } from "documonster/excel";
 *
 * const parts = Cell.getDateParts(sheet, "A1");
 * // { year: 2024, month: 1, day: 15, hour: 0, minute: 0, second: 0, millisecond: 0 }
 * ```
 */
export function getDateParts(ws: WorksheetData, addr: string): ExcelDateTimeParts | undefined;
/** Read a date cell as calendar fields, by 1-based (row, col). */
export function getDateParts(
  ws: WorksheetData,
  row: number,
  col: number
): ExcelDateTimeParts | undefined;
export function getDateParts(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): ExcelDateTimeParts | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellGetDateParts(cell) : undefined;
}

/**
 * Write a date cell from calendar fields, saying which of the three kinds of date cell it is.
 *
 * Missing fields default to zero, so `{ year: 2024, month: 1, day: 15 }` is a date and `{ hour: 9, minute: 30 }`
 * is a time. `kind` decides the default number format and is inferred from the fields present when omitted — it
 * is worth passing explicitly for a date-time whose time happens to be midnight, which is otherwise
 * indistinguishable from a date.
 *
 * This is the runtime-independent half of the Temporal surface: it needs nothing but the fields.
 *
 * @example
 * ```ts
 * import { Cell } from "documonster/excel";
 *
 * Cell.setDateParts(sheet, "A1", { year: 2024, month: 1, day: 15 });
 * Cell.setDateParts(sheet, "A2", { hour: 9, minute: 30 }, "time");
 * ```
 */
export function setDateParts(
  ws: WorksheetData,
  addr: string,
  parts: Partial<ExcelDateTimeParts>,
  kind?: TemporalKind
): void;
/** Write a date cell from calendar fields, by 1-based (row, col). */
export function setDateParts(
  ws: WorksheetData,
  row: number,
  col: number,
  parts: Partial<ExcelDateTimeParts>,
  kind?: TemporalKind
): void;
export function setDateParts(
  ws: WorksheetData,
  addr: string | number,
  partsOrCol: Partial<ExcelDateTimeParts> | number,
  partsOrKind?: Partial<ExcelDateTimeParts> | TemporalKind,
  kind?: TemporalKind
): void {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  // `arguments.length` decides the form, as for every other setter: a `(ws, "A1", 5, parts)` call used to take `5` as
  // a column and address `EA1`.
  const positional = arguments.length >= 4 && typeof partsOrCol === "number";
  const parts = (positional ? partsOrKind : partsOrCol) as Partial<ExcelDateTimeParts>;
  const resolvedKind = (positional ? kind : partsOrKind) as TemporalKind | undefined;
  // Checked before the cell is looked up, so a refused value leaves no new empty cell behind.
  const check = argumentChecks("Cell.setDateParts");
  check.required("parts", parts);
  check.record("parts", parts);
  check.oneOf("kind", resolvedKind, ["date", "time", "dateTime"]);
  validateDateParts(parts, getSheetWorkbook(ws).properties.date1904 === true);
  const cell = positional ? getCell(ws, rowNumber(addr), partsOrCol) : getCell(ws, addr as string);
  cellSetDateParts(cell, parts, resolvedKind);
}

/**
 * What kind of date cell this is, or `undefined` for a cell that holds no date.
 *
 * Decided by the cell's number format, which is the only place a spreadsheet records the distinction: the
 * stored value is a serial, and a serial does not say whether the whole-day part or the fraction is the point.
 *
 * - `"date"`, `"time"`, `"dateTime"` — the format settles it.
 * - `"unknown"` — the cell holds a date, but its format does not say which kind. An unformatted cell, a
 *   `General` one, or one wearing a plain number format.
 * - `"duration"` — an elapsed-time format such as `[h]:mm:ss`, where the serial is a length of time rather
 *   than a moment. {@link getTemporal} refuses these; {@link getDateParts} does not.
 */
export function getDateKind(ws: WorksheetData, addr: string): DateFormatKind | undefined;
/** What kind of date cell this is, by 1-based (row, col). */
export function getDateKind(
  ws: WorksheetData,
  row: number,
  col: number
): DateFormatKind | undefined;
export function getDateKind(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): DateFormatKind | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellDateKind(cell) : undefined;
}

/**
 * Read a date cell as a Temporal `Plain*` value, chosen by the cell's number format.
 *
 * A `PlainDate` for a date-formatted cell, a `PlainTime` for a time-formatted one, a `PlainDateTime` for one
 * carrying both, or for one whose format does not say. Returns `undefined` for a cell holding no date.
 *
 * **Refuses an elapsed-time format** such as `[h]:mm:ss`: that serial is a length of time, not a calendar
 * value, and there is no `Plain*` that means it. Pass `kind` to read it as a civil value regardless.
 *
 * **Refuses Excel's fictitious 1900-02-29** (serial 60), which the ISO calendar has no room for.
 * {@link getDateParts} reports it as plain fields.
 *
 * **Throws where Temporal is absent** — Node below 26, Safari, Chrome below 144 — rather than quietly
 * returning a `Date` instead, because a return type that depends on the runtime is worse than a refusal. This
 * package adds no polyfill; the zero-dependency rule forbids one. {@link getDateParts} is the same value
 * everywhere.
 *
 * @param kind - Override the format's verdict.
 *
 * @example
 * ```ts
 * import { Cell } from "documonster/excel";
 *
 * const when = Cell.getTemporal(sheet, "A1"); // Temporal.PlainDate 2024-01-15
 * ```
 */
export function getTemporal(
  ws: WorksheetData,
  addr: string,
  kind?: TemporalKind
): TemporalPlainValue | undefined;
/** Read a date cell as a Temporal `Plain*` value, by 1-based (row, col). */
export function getTemporal(
  ws: WorksheetData,
  row: number,
  col: number,
  kind?: TemporalKind
): TemporalPlainValue | undefined;
export function getTemporal(
  ws: WorksheetData,
  addr: string | number,
  colOrKind?: number | TemporalKind,
  kind?: TemporalKind
): TemporalPlainValue | undefined {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  const cell =
    typeof colOrKind === "number"
      ? locateCell(ws, rowNumber(addr), colOrKind)
      : locateCell(ws, addr);
  return cell ? cellGetTemporal(cell, typeof colOrKind === "number" ? kind : colOrKind) : undefined;
}

/** Read a cell's text by `"A1"` address. */
export function getText(ws: WorksheetData, addr: string): string;
/** Read a cell's text by 1-based (row, col). */
export function getText(ws: WorksheetData, row: number, col: number): string;
export function getText(ws: WorksheetData, addr: string | number, col?: number): string {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellText(cell) : "";
}
/** Read a cell's display text by `"A1"` address. */
export function getDisplayText(ws: WorksheetData, addr: string): string;
/** Read a cell's display text by 1-based (row, col). */
export function getDisplayText(ws: WorksheetData, row: number, col: number): string;
export function getDisplayText(ws: WorksheetData, addr: string | number, col?: number): string {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellDisplayText(cell) : "";
}
/** Read a cell's value kind by `"A1"` address. */
export function getType(ws: WorksheetData, addr: string): ValueType;
/** Read a cell's value kind by 1-based (row, col). */
export function getType(ws: WorksheetData, row: number, col: number): ValueType;
export function getType(ws: WorksheetData, addr: string | number, col?: number): ValueType {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellType(cell) : ValueType.Null;
}
/** Read a cell's effective value kind by `"A1"` address. */
export function getEffectiveType(ws: WorksheetData, addr: string): ValueType;
/** Read a cell's effective value kind by 1-based (row, col). */
export function getEffectiveType(ws: WorksheetData, row: number, col: number): ValueType;
export function getEffectiveType(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): ValueType {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellEffectiveType(cell) : ValueType.Null;
}

// --- formula ---

/**
 * Read a cell's formula by `"A1"` address.
 *
 * The text is what the cell holds. For a formula read from a file that is the stored form, with the prefixes the file
 * carries (`_xlfn.XLOOKUP(…)`, `_xlpm.x`, `_xleta.SUM`) — including for a formula that was set bare and then written.
 */
export function getFormula(ws: WorksheetData, addr: string): string | undefined;
/** Read a cell's formula by 1-based (row, col). */
export function getFormula(ws: WorksheetData, row: number, col: number): string | undefined;
export function getFormula(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): string | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellFormula(cell) : undefined;
}
/** Read a formula cell's cached result by `"A1"` address. */
export function getResult(ws: WorksheetData, addr: string): FormulaResult | undefined;
/** Read a formula cell's cached result by 1-based (row, col). */
export function getResult(ws: WorksheetData, row: number, col: number): FormulaResult | undefined;
export function getResult(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): FormulaResult | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellResult(cell) : undefined;
}

// --- style ---

/**
 * The live style of the cell at `"A1"`, for editing in place — it creates the cell when there is none.
 * To read without creating anything, use {@link findStyle}.
 */
export function getStyle(ws: WorksheetData, addr: string): Partial<Style>;
/** Read a cell's style by 1-based (row, col). */
export function getStyle(ws: WorksheetData, row: number, col: number): Partial<Style>;
export function getStyle(ws: WorksheetData, addr: string | number, col?: number): Partial<Style> {
  return cellOwnStyle(target(ws, addr, col, arguments.length));
}
/**
 * A copy of the style of the cell at `"A1"`, or `undefined` when the sheet has no cell there — the
 * read-only counterpart of {@link getStyle}, which hands out the live style for editing and so creates
 * the cell. This creates nothing, and editing what it returns changes nothing.
 */
export function findStyle(ws: WorksheetData, addr: string): Partial<Style> | undefined;
/** {@link findStyle} by 1-based (row, col). */
export function findStyle(ws: WorksheetData, row: number, col: number): Partial<Style> | undefined;
export function findStyle(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): Partial<Style> | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? (copyStyle(cellGetStyle(cell)) as Partial<Style>) : undefined;
}
/** Merge a partial style into the cell at "A1" address. */
export function setStyle(ws: WorksheetData, addr: string, style: Partial<Style>): void;
/** Merge a partial style into the cell at 1-based (row, col). */
export function setStyle(ws: WorksheetData, row: number, col: number, style: Partial<Style>): void;
export function setStyle(
  ws: WorksheetData,
  addr: string | number,
  styleOrCol: Partial<Style> | number,
  style?: Partial<Style>
): void {
  const [cell, resolved] = targetWithValue<Partial<Style>>(
    ws,
    addr,
    styleOrCol,
    style,
    arguments.length,
    value => checkStyle("Cell.setStyle", value)
  );
  cellSetStyle(cell, resolved);
}

/**
 * Apply a workbook-level named cell style (e.g. "Heading 1") to the cell at
 * the "A1" address. The style must first be defined with
 * `Workbook.addCellStyle`; applying an unknown name throws. To set a raw
 * `styleName` without this check, use `Cell.setStyle(ws, addr, { styleName })`.
 */
export function applyStyle(ws: WorksheetData, addr: string, name: string): void;
/** Apply a named cell style to the cell at 1-based (row, col). */
export function applyStyle(ws: WorksheetData, row: number, col: number, name: string): void;
export function applyStyle(
  ws: WorksheetData,
  addr: string | number,
  nameOrCol: string | number,
  name?: string
): void {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
  const styleName = (arguments.length >= 4 ? name : nameOrCol) as string;
  argumentChecks("Cell.applyStyle").required("name", styleName);
  argumentChecks("Cell.applyStyle").string("name", styleName, { nonEmpty: true });
  // Found in any case, as Excel finds it; the cell records the style's own spelling.
  const named = getCellStyle(getSheetWorkbook(ws), styleName);
  if (!named) {
    throw new ExcelArgumentError(
      "name",
      `Named cell style "${styleName}" is not defined. Define it first with Workbook.addCellStyle().`
    );
  }
  if (arguments.length >= 4) {
    cellSetStyle(getCell(ws, rowNumber(addr), nameOrCol as number), { styleName: named.name });
    return;
  }
  cellSetStyle(getCell(ws, addr), { styleName: named.name });
}

/** @deprecated Use {@link applyStyle}. Removed in 1.0. */
export const applyCellStyle = applyStyle;

// --- merge ---

/** Test whether a cell is merged by `"A1"` address. */
export function isMerged(ws: WorksheetData, addr: string): boolean;
/** Test whether a cell is merged by 1-based (row, col). */
export function isMerged(ws: WorksheetData, row: number, col: number): boolean;
export function isMerged(ws: WorksheetData, addr: string | number, col?: number): boolean {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellIsMerged(cell) : false;
}
/** The address of the merge's master (top-left) cell, or the cell's own address when it is not merged. */
export function getMergeMasterAddress(ws: WorksheetData, addr: string): string;
/** The master address of the merge covering 1-based (row, col). */
export function getMergeMasterAddress(ws: WorksheetData, row: number, col: number): string;
export function getMergeMasterAddress(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): string {
  const at = addressOf(ws, addr, col, arguments.length);
  const cell = locateCell(ws, at.address);
  return cell ? cellMaster(cell).address : at.address;
}

/** @deprecated Use {@link getMergeMasterAddress}. Removed in 1.0. */
export function getMergeMaster(ws: WorksheetData, addr: string): CellData;
/** @deprecated Use {@link getMergeMasterAddress}. Removed in 1.0. */
export function getMergeMaster(ws: WorksheetData, row: number, col: number): CellData;
export function getMergeMaster(ws: WorksheetData, addr: string | number, col?: number): CellData {
  checkWorksheet(ws);
  // A merged cell always exists, so an absent one is its own master: preview it rather than create it.
  return cellMaster(arguments.length >= 3 ? previewCell(ws, addr, col) : previewCell(ws, addr));
}

// --- hyperlink ---

/** Read a cell's hyperlink by `"A1"` address. */
export function getHyperlink(ws: WorksheetData, addr: string): string | undefined;
/** Read a cell's hyperlink by 1-based (row, col). */
export function getHyperlink(ws: WorksheetData, row: number, col: number): string | undefined;
export function getHyperlink(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): string | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellHyperlink(cell) : undefined;
}

// --- note ---

/** Read a cell's note by `"A1"` address. */
export function getNote(ws: WorksheetData, addr: string): string | NoteConfig | undefined;
/** Read a cell's note by 1-based (row, col). */
export function getNote(
  ws: WorksheetData,
  row: number,
  col: number
): string | NoteConfig | undefined;
export function getNote(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): string | NoteConfig | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellNote(cell) : undefined;
}
/** Set a cell's note by `"A1"` address. */
export function setNote(ws: WorksheetData, addr: string, note: string | NoteConfig): void;
/** Set a cell's note by 1-based (row, col). */
export function setNote(
  ws: WorksheetData,
  row: number,
  col: number,
  note: string | NoteConfig
): void;
export function setNote(
  ws: WorksheetData,
  addr: string | number,
  noteOrCol: string | NoteConfig | number,
  note?: string | NoteConfig
): void {
  const [cell, resolved] = targetWithValue<string | NoteConfig>(
    ws,
    addr,
    noteOrCol,
    note,
    arguments.length,
    value => {
      if (typeof value !== "string" && (typeof value !== "object" || value === null)) {
        argumentChecks("Cell.setNote").fail("note", "a string or a note configuration", value);
      }
      checkNote("Cell.setNote", "note", value);
    }
  );
  cellSetNote(cell, resolved);
}

// --- defined names ---

/** Read the defined names covering a cell by `"A1"` address. */
export function getNames(ws: WorksheetData, addr: string): string[];
/** Read the defined names covering a cell by 1-based (row, col). */
export function getNames(ws: WorksheetData, row: number, col: number): string[];
export function getNames(ws: WorksheetData, addr: string | number, col?: number): string[] {
  // Defined names live on the workbook, keyed by address; the cell itself holds none.
  const address = addressOf(ws, addr, col, arguments.length);
  return definedNamesGetNamesEx(getSheetWorkbook(ws)._definedNames, address);
}
/** Add a defined name to a cell by `"A1"` address. */
export function addName(ws: WorksheetData, addr: string, name: string): void;
/** Add a defined name to a cell by 1-based (row, col). */
export function addName(ws: WorksheetData, row: number, col: number, name: string): void;
export function addName(
  ws: WorksheetData,
  addr: string | number,
  nameOrCol: string | number,
  name?: string
): void {
  const [cell, resolved] = targetWithValue<string>(
    ws,
    addr,
    nameOrCol,
    name,
    arguments.length,
    value => checkCellName(ws, "Cell.addName", value)
  );
  cellAddName(cell, resolved);
}
/**
 * Remove a defined name from the cell at `"A1"`. Returns `true` if the cell carried it. Creates nothing: the
 * names live on the workbook, keyed by address.
 */
export function removeName(ws: WorksheetData, addr: string, name: string): boolean;
/** Remove a defined name from the cell at 1-based (row, col). Returns `true` if the cell carried it. */
export function removeName(ws: WorksheetData, row: number, col: number, name: string): boolean;
export function removeName(
  ws: WorksheetData,
  addr: string | number,
  nameOrCol: string | number,
  name?: string
): boolean {
  // The names live on the workbook, keyed by sheet name: through a removed sheet's handle this would edit the names
  // of whichever sheet has taken the name since.
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  const [address, resolved] =
    arguments.length >= 4
      ? [resolveAddress(ws, rowNumber(addr), nameOrCol as number), name as string]
      : [resolveAddress(ws, addr), nameOrCol as string];
  argumentChecks("Cell.removeName").string("name", resolved);
  return definedNamesRemoveEx(getSheetWorkbook(ws)._definedNames, address, resolved);
}
/** Replace a cell's defined names with one by `"A1"` address. */
export function setName(ws: WorksheetData, addr: string, name: string): void;
/** Replace a cell's defined names with one by 1-based (row, col). */
export function setName(ws: WorksheetData, row: number, col: number, name: string): void;
export function setName(
  ws: WorksheetData,
  addr: string | number,
  nameOrCol: string | number,
  name?: string
): void {
  const [cell, resolved] = targetWithValue<string>(
    ws,
    addr,
    nameOrCol,
    name,
    arguments.length,
    value => checkCellName(ws, "Cell.setName", value)
  );
  cellSetName(cell, resolved);
}
/** Replace a cell's defined names by `"A1"` address. */
export function setNames(ws: WorksheetData, addr: string, names: readonly string[]): void;
/** Replace a cell's defined names by 1-based (row, col). */
export function setNames(
  ws: WorksheetData,
  row: number,
  col: number,
  names: readonly string[]
): void;
export function setNames(
  ws: WorksheetData,
  addr: string | number,
  namesOrCol: readonly string[] | number,
  names?: readonly string[]
): void {
  const [cell, resolved] = targetWithValue<readonly string[]>(
    ws,
    addr,
    namesOrCol,
    names,
    arguments.length,
    value => {
      if (!Array.isArray(value)) {
        argumentChecks("Cell.setNames").fail("names", "an array of names", value);
      }
      value.forEach(n => checkCellName(ws, "Cell.setNames", n));
    }
  );
  cellSetNames(cell, [...resolved]);
}
/** Remove every defined name from a cell by `"A1"` address. */
export function removeAllNames(ws: WorksheetData, addr: string): void;
/** Remove every defined name from a cell by 1-based (row, col). */
export function removeAllNames(ws: WorksheetData, row: number, col: number): void;
export function removeAllNames(ws: WorksheetData, addr: string | number, col?: number): void {
  // See `removeName`: the names are keyed by sheet name, so a removed sheet's handle is refused.
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  definedNamesRemoveAllNames(
    getSheetWorkbook(ws)._definedNames,
    addressOf(ws, addr, col, arguments.length)
  );
}

// --- data validation ---

/** A copy of the validation rule covering a cell, by `"A1"` address, or `undefined`. */
export function getValidation(ws: WorksheetData, addr: string): DataValidationRule | undefined;
/** Read a cell's validation rule by 1-based (row, col). */
export function getValidation(
  ws: WorksheetData,
  row: number,
  col: number
): DataValidationRule | undefined;
export function getValidation(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): DataValidationRule | undefined {
  // Validation rules live on the sheet, keyed by address; the cell itself holds none. A copy, as the other facet
  // readers return: editing it changes nothing until it is set again.
  const address = addressOf(ws, addr, col, arguments.length);
  const rule = dataValidationFind(ws.dataValidations, address.address);
  return rule === undefined ? undefined : structuredClone(rule);
}
/** Set a cell's validation rule by `"A1"` address. The sheet keeps a copy of `value`. */
export function setValidation(ws: WorksheetData, addr: string, value: DataValidationRule): void;
/** Set a cell's validation rule by 1-based (row, col). */
export function setValidation(
  ws: WorksheetData,
  row: number,
  col: number,
  value: DataValidationRule
): void;
export function setValidation(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: DataValidationRule | number,
  value?: DataValidationRule
): void {
  const [cell, resolved] = targetWithValue<DataValidationRule>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    rule => checkDataValidationRule("Cell.setValidation", rule)
  );
  // A copy: the caller's object used to be the stored rule, so one rule given to two cells — or edited afterwards —
  // changed what every one of them wrote.
  cellSetDataValidation(cell, resolved === undefined ? resolved : structuredClone(resolved));
}

// --- model (advanced / round-trip) ---

/**
 * Read a cell's round-trip model by `"A1"` address. For an existing cell it is the cell's own, live model — edit the
 * cell through the setters, or apply a changed copy with {@link setModel}. Creates nothing: an absent cell reports
 * what it would inherit.
 */
export function getModel(ws: WorksheetData, addr: string): CellModel;
/** Read a cell's round-trip model by 1-based (row, col). */
export function getModel(ws: WorksheetData, row: number, col: number): CellModel;
export function getModel(ws: WorksheetData, addr: string | number, col?: number): CellModel {
  return cellGetModel(own(ws, addr, col, arguments.length));
}
/**
 * Apply a round-trip cell model by `"A1"` address.
 *
 * @throws {ExcelArgumentError} when the model is for another address, has an unknown `type`, or holds a value of
 * another kind than its `type` or one no workbook can hold (`NaN`, an invalid `Date`).
 */
export function setModel(ws: WorksheetData, addr: string, model: CellModel): void;
/** Apply a round-trip cell model by 1-based (row, col). */
export function setModel(ws: WorksheetData, row: number, col: number, model: CellModel): void;
export function setModel(
  ws: WorksheetData,
  addr: string | number,
  modelOrCol: CellModel | number,
  model?: CellModel
): void {
  // The address is resolved first, so a model for another cell is refused before this one is created: it was
  // written into this cell's `<row>` under its own address, and the file could not be read back.
  const { address } = addressOf(ws, addr, modelOrCol as number, arguments.length >= 4 ? 3 : 2);
  const [cell, resolved] = targetWithValue<CellModel>(
    ws,
    addr,
    modelOrCol,
    model,
    arguments.length,
    value => checkCellModel("Cell.setModel", "model", value, address)
  );
  cellSetModel(cell, resolved);
}

// --- individual style facets (getters + setters) ---

/** Read a cell's font by `"A1"` address. */
export function getFont(ws: WorksheetData, addr: string): Partial<Font> | undefined;
/** Read a cell's font by 1-based (row, col). */
export function getFont(ws: WorksheetData, row: number, col: number): Partial<Font> | undefined;
export function getFont(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): Partial<Font> | undefined {
  return cellFont(own(ws, addr, col, arguments.length));
}
/** Set a cell's font by `"A1"` address. */
export function setFont(ws: WorksheetData, addr: string, value: Partial<Font> | undefined): void;
/** Set a cell's font by 1-based (row, col). */
export function setFont(
  ws: WorksheetData,
  row: number,
  col: number,
  value: Partial<Font> | undefined
): void;
export function setFont(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: Partial<Font> | undefined | number,
  value?: Partial<Font>
): void {
  const [cell, resolved] = targetWithValue<Partial<Font> | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    value => checkStyleFacet("Cell.setFont", "font", value, "value")
  );
  cellSetFont(cell, resolved);
}
/** Read a cell's number format by `"A1"` address. */
export function getNumFmt(ws: WorksheetData, addr: string): string | NumFmt | undefined;
/** Read a cell's number format by 1-based (row, col). */
export function getNumFmt(ws: WorksheetData, row: number, col: number): string | NumFmt | undefined;
export function getNumFmt(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): string | NumFmt | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  if (cell) {
    return cellNumFmt(cell);
  }
  const { row, col: column } = addressOf(ws, addr, col, arguments.length);
  return inheritedNumFmt(ws, row, column);
}
/** Set a cell's number format by `"A1"` address. */
export function setNumFmt(
  ws: WorksheetData,
  addr: string,
  value: string | NumFmt | undefined
): void;
/** Set a cell's number format by 1-based (row, col). */
export function setNumFmt(
  ws: WorksheetData,
  row: number,
  col: number,
  value: string | NumFmt | undefined
): void;
export function setNumFmt(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: string | NumFmt | undefined | number,
  value?: string | NumFmt
): void {
  const [cell, resolved] = targetWithValue<string | NumFmt | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    value => checkStyleFacet("Cell.setNumFmt", "numFmt", value, "value")
  );
  cellSetNumFmt(cell, resolved);
}
/** Read a cell's alignment by `"A1"` address. */
export function getAlignment(ws: WorksheetData, addr: string): Partial<Alignment> | undefined;
/** Read a cell's alignment by 1-based (row, col). */
export function getAlignment(
  ws: WorksheetData,
  row: number,
  col: number
): Partial<Alignment> | undefined;
export function getAlignment(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): Partial<Alignment> | undefined {
  return cellAlignment(own(ws, addr, col, arguments.length));
}
/** Set a cell's alignment by `"A1"` address. */
export function setAlignment(
  ws: WorksheetData,
  addr: string,
  value: Partial<Alignment> | undefined
): void;
/** Set a cell's alignment by 1-based (row, col). */
export function setAlignment(
  ws: WorksheetData,
  row: number,
  col: number,
  value: Partial<Alignment> | undefined
): void;
export function setAlignment(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: Partial<Alignment> | undefined | number,
  value?: Partial<Alignment>
): void {
  const [cell, resolved] = targetWithValue<Partial<Alignment> | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    value => checkStyleFacet("Cell.setAlignment", "alignment", value, "value")
  );
  cellSetAlignment(cell, resolved);
}
/** Read a cell's borders by `"A1"` address. */
export function getBorder(ws: WorksheetData, addr: string): Partial<Borders> | undefined;
/** Read a cell's borders by 1-based (row, col). */
export function getBorder(
  ws: WorksheetData,
  row: number,
  col: number
): Partial<Borders> | undefined;
export function getBorder(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): Partial<Borders> | undefined {
  return cellBorder(own(ws, addr, col, arguments.length));
}
/** Set a cell's borders by `"A1"` address. */
export function setBorder(
  ws: WorksheetData,
  addr: string,
  value: Partial<Borders> | undefined
): void;
/** Set a cell's borders by 1-based (row, col). */
export function setBorder(
  ws: WorksheetData,
  row: number,
  col: number,
  value: Partial<Borders> | undefined
): void;
export function setBorder(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: Partial<Borders> | undefined | number,
  value?: Partial<Borders>
): void {
  const [cell, resolved] = targetWithValue<Partial<Borders> | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    value => checkStyleFacet("Cell.setBorder", "border", value, "value")
  );
  cellSetBorder(cell, resolved);
}
/** Read a cell's fill by `"A1"` address. */
export function getFill(ws: WorksheetData, addr: string): Fill | undefined;
/** Read a cell's fill by 1-based (row, col). */
export function getFill(ws: WorksheetData, row: number, col: number): Fill | undefined;
export function getFill(ws: WorksheetData, addr: string | number, col?: number): Fill | undefined {
  return cellFill(own(ws, addr, col, arguments.length));
}
/** Set a cell's fill by `"A1"` address. */
export function setFill(ws: WorksheetData, addr: string, value: Fill | undefined): void;
/** Set a cell's fill by 1-based (row, col). */
export function setFill(ws: WorksheetData, row: number, col: number, value: Fill | undefined): void;
export function setFill(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: Fill | undefined | number,
  value?: Fill
): void {
  const [cell, resolved] = targetWithValue<Fill | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    value => checkStyleFacet("Cell.setFill", "fill", value, "value")
  );
  cellSetFill(cell, resolved);
}
/** Read a cell's protection by `"A1"` address. */
export function getProtection(ws: WorksheetData, addr: string): Partial<Protection> | undefined;
/** Read a cell's protection by 1-based (row, col). */
export function getProtection(
  ws: WorksheetData,
  row: number,
  col: number
): Partial<Protection> | undefined;
export function getProtection(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): Partial<Protection> | undefined {
  return cellProtection(own(ws, addr, col, arguments.length));
}
/** Set a cell's protection by `"A1"` address. */
export function setProtection(
  ws: WorksheetData,
  addr: string,
  value: Partial<Protection> | undefined
): void;
/** Set a cell's protection by 1-based (row, col). */
export function setProtection(
  ws: WorksheetData,
  row: number,
  col: number,
  value: Partial<Protection> | undefined
): void;
export function setProtection(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: Partial<Protection> | undefined | number,
  value?: Partial<Protection>
): void {
  const [cell, resolved] = targetWithValue<Partial<Protection> | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    arguments.length,
    value => checkStyleFacet("Cell.setProtection", "protection", value, "value")
  );
  cellSetProtection(cell, resolved);
}

// --- comment (author-bearing note) ---

/** Read a cell's comment by `"A1"` address. */
export function getComment(ws: WorksheetData, addr: string): NoteData | undefined;
/** Read a cell's comment by 1-based (row, col). */
export function getComment(ws: WorksheetData, row: number, col: number): NoteData | undefined;
export function getComment(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): NoteData | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell ? cellComment(cell) : undefined;
}
/** Set a cell's comment by `"A1"` address. */
export function setComment(
  ws: WorksheetData,
  addr: string,
  comment: NoteData | NoteConfig | undefined
): void;
/** Set a cell's comment by 1-based (row, col). */
export function setComment(
  ws: WorksheetData,
  row: number,
  col: number,
  comment: NoteData | NoteConfig | undefined
): void;
export function setComment(
  ws: WorksheetData,
  addr: string | number,
  commentOrCol: NoteData | NoteConfig | undefined | number,
  comment?: NoteData | NoteConfig
): void {
  const [cell, resolved] = targetWithValue<NoteData | NoteConfig | undefined>(
    ws,
    addr,
    commentOrCol,
    comment,
    arguments.length,
    value => {
      if (value !== undefined && (typeof value !== "object" || value === null)) {
        argumentChecks("Cell.setComment").fail(
          "comment",
          "a note, a note configuration or undefined",
          value
        );
      }
      if (isNoteData(value)) {
        checkNote("Cell.setComment", "comment.note", value.note);
        if (value.author !== undefined && value.author !== null) {
          argumentChecks("Cell.setComment").string("comment.author", value.author);
        }
      } else {
        checkNote("Cell.setComment", "comment", value);
      }
    }
  );
  cellSetComment(cell, resolved);
}

// --- formula result / full address ---

/**
 * Set a formula cell's cached result by `"A1"` address — what a reader shows until the workbook is recalculated.
 *
 * @throws {ExcelArgumentError} when the cell holds no formula (passing `undefined` there is allowed and does nothing).
 */
export function setResult(ws: WorksheetData, addr: string, value: FormulaResult | undefined): void;
/** Set a formula cell's cached result by 1-based (row, col). */
export function setResult(
  ws: WorksheetData,
  row: number,
  col: number,
  value: FormulaResult | undefined
): void;
export function setResult(
  ws: WorksheetData,
  addr: string | number,
  valueOrCol: FormulaResult | undefined | number,
  value?: FormulaResult
): void {
  const argc = arguments.length;
  const [cell, resolved] = targetWithValue<FormulaResult | undefined>(
    ws,
    addr,
    valueOrCol,
    value,
    argc,
    result => {
      checkFormulaResult(result, { where: "Cell.setResult", field: "value" });
      // A cached result belongs to a formula. On any other cell it was dropped without a word.
      const existing =
        argc >= 4 ? peek(ws, addr, valueOrCol as number, 3) : peek(ws, addr, undefined, 2);
      if (result !== undefined && (!existing || cellType(existing) !== ValueType.Formula)) {
        argumentChecks("Cell.setResult").fail(
          "value",
          "set on a cell that holds a formula",
          result
        );
      }
    }
  );
  cellSetResult(cell, resolved);
}
/** The cell's fully-qualified address by `"A1"` address. */
export function getFullAddress(ws: WorksheetData, addr: string): DecodedAddress;
/** The cell's fully-qualified address by 1-based (row, col). */
export function getFullAddress(ws: WorksheetData, row: number, col: number): DecodedAddress;
export function getFullAddress(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): DecodedAddress {
  return addressOf(ws, addr, col, arguments.length);
}

// --- existence ---

/**
 * Whether the sheet has a cell at the address — **without creating it**.
 *
 * The value readers — `getValue`, `getText`, `getFormula`, `getNumFmt`, … — do
 * not create anything either, and answer for a missing cell as for an empty one.
 * Use `has` when the question is whether a cell is there at all — probing a
 * sparse sheet, walking the ghost cells of a spilled dynamic array, or telling an
 * absent cell from an empty one. Of the readers only `getStyle` creates the cell,
 * because it hands out the live style for editing in place; the facet readers
 * (`getFont`, …) and `getModel` answer for an absent cell with what it would
 * inherit, and `findStyle` reads a cell's style without creating anything.
 *
 * @example
 * ```typescript
 * import { Cell, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Sheet1");
 * const value = Cell.has(ws, "B7") ? Cell.getValue(ws, "B7") : null;
 * ```
 */
export function has(ws: WorksheetData, addr: string): boolean;
/** Whether the sheet has a cell at 1-based (row, col), without creating it. */
export function has(ws: WorksheetData, row: number, col: number): boolean;
export function has(ws: WorksheetData, addr: string | number, col?: number): boolean {
  // `argc` decides the overload, as in `target`.
  return peek(ws, addr, col, arguments.length) !== undefined;
}

/**
 * A read-only view of the cell at `addr` — value, number format, text, effective
 * type, font and alignment — or `undefined` when no cell exists there. Never
 * creates the cell.
 *
 * The view reads the cell live and copies nothing, so it is the allocation-free
 * way to scan a large sheet. Its `font` / `alignment` are shared and frozen; change
 * them through {@link setFont} / {@link setAlignment}.
 *
 * @example
 * ```typescript
 * import { Cell, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Data");
 * Cell.setValue(ws, "A1", 42);
 * Cell.getView(ws, "A1")?.text; // "42"
 * Cell.getView(ws, "B9"); // undefined — and B9 is still absent
 * ```
 */
export function getView(ws: WorksheetData, addr: string): CellView | undefined;
/** A read-only view of the cell at 1-based (row, col), without creating it. */
export function getView(ws: WorksheetData, row: number, col: number): CellView | undefined;
export function getView(
  ws: WorksheetData,
  addr: string | number,
  col?: number
): CellView | undefined {
  const cell = peek(ws, addr, col, arguments.length);
  return cell && cellView(cell);
}

// --- deprecated cell handles ---

/**
 * @deprecated Use {@link has} to test for a cell, {@link getView} to read one
 * without creating it, and address it by position. Removed in 1.0.
 */
export function find(ws: WorksheetData, addr: string): CellData | undefined;
/** @deprecated Use {@link has} / {@link getView}. Removed in 1.0. */
export function find(ws: WorksheetData, row: number, col: number): CellData | undefined;
export function find(ws: WorksheetData, addr: string | number, col?: number): CellData | undefined {
  return peek(ws, addr, col, arguments.length);
}

/**
 * @deprecated The document model is addressed by position; use {@link getView}
 * (`Cell.getView(ws, addr)`), or {@link getValue} and friends. For a streaming handle use `Stream.getCellView`.
 * Removed in 1.0.
 */
export function view(cell: CellData): CellView {
  if ((cell as { _value?: unknown } | null | undefined)?._value === undefined) {
    argumentChecks("Cell.view").fail("cell", "a cell", cell);
  }
  return cellView(cell);
}

// --- types ---

/**
 * A value read out of a cell — what {@link getValue} returns, and the element
 * type of the matrix `Range.getValues` returns.
 */
export type Value = CellValueType;
