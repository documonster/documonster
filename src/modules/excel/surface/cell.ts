/**
 * `Cell` namespace surface — public `(ws, addr, …)` cell operations.
 *
 * Consumed as `import { Cell } from "documonster/excel"` → `Cell.setValue(ws, "A1", 42)`.
 * Each function resolves the cell handle via `getCell(ws, addr)` and delegates
 * to the internal handle-level helpers. Consumers never hold a `CellData`.
 *
 * The `addr` parameter is either an `"A1"` string or a 1-based `{ row, col }`
 * position (`Cell.setValue(ws, { row: 1, col: 1 }, 42)`). The `(ws, row, col, …)`
 * form is deprecated; it is told apart by its numeric row, never by argument count.
 *
 * This is a flat-named-export module; `excel/index.ts` re-exports it via
 * `export * as Cell`, which tree-shakes per-member on rolldown / rspack.
 */
import type { CellData, CellModel, CellView, FormulaResult, NoteConfig } from "@excel/core/cell";
import {
  checkCellValue,
  checkFormulaResult,
  checkSharedFormulaAt,
  cellAddName,
  cellDateKind,
  cellGetDateParts,
  cellGetTemporal,
  cellSetDateParts,
  validateDateParts,
  cellComment,
  cellDisplayText,
  cellEffectiveType,
  cellFormula,
  cellGetModel,
  cellGetStyle,
  cellOwnStyle,
  cellReadValue,
  cellHyperlink,
  cellIsMerged,
  cellMaster,
  cellNote,
  cellNumFmt,
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
  cellView,
  ownValue
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
import { ExcelArgumentError, ExcelInvalidAddressError } from "@excel/errors";
import { checkDataValidationRule, checkNote } from "@excel/surface/argument-checks";
import { checkCellModel } from "@excel/surface/model-checks";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import {
  asStreamWorksheet,
  checkStreamRowsHeld,
  checkStreamRowsOpen
} from "@excel/surface/stream-targets";
import type {
  Alignment,
  Borders,
  CellPosition,
  CellValue,
  CellValueInput,
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
import { copyPlain } from "@excel/utils/copy-plain";
import { copyStyle, copyStyleFacet } from "@excel/utils/copy-style";
import { checkStyle, checkStyleFacet } from "@excel/utils/style-checks";
import type { ExcelDateTimeParts } from "@utils/excel-serial";
import { isRecordArgument } from "@utils/object";

/** @deprecated Use `Worksheet.Handle`. Removed in 1.0. */
export type Sheet = WorksheetData;

const MAX_ROW = 1048576;
const MAX_COL = 16384;

/**
 * The cell named by a function's address arguments, as the `(r, c)` pair the core resolvers take.
 *
 * The address is `"A1"` text or a `{ row, col }` position. The deprecated `(row, col)` form is told apart by the type
 * of its first argument — a number — never by how many arguments were passed, so a setter given too few arguments
 * cannot take its column for its value.
 *
 * `stray` is the argument after an `"A1"` address or a position, which the deprecated form reads as a column. Handed
 * `(ws, "A1", 5)`, the resolver once built the address `n2l(5) + "A1"` — `EA1`, column 131 — so it is refused.
 */
function ref(
  where: string,
  addr: string | CellPosition | number,
  col: number | undefined,
  stray: unknown
): [string | number, number | undefined] {
  if (typeof addr === "number") {
    return [addr, col];
  }
  if (stray !== undefined) {
    throw new ExcelInvalidAddressError(
      String(addr),
      'the (row, col) form takes a row number; an "A1" address takes no column',
      { where }
    );
  }
  if (isRecordArgument(addr)) {
    checkPosition(where, addr);
    return [addr.row, addr.col];
  }
  return [addr, undefined];
}

/** A `{ row, col }` position names a cell inside A1:XFD1048576 with whole numbers. */
function checkPosition(where: string, addr: CellPosition): void {
  const bounds = [
    ["row", MAX_ROW],
    ["col", MAX_COL]
  ] as const;
  for (const [key, max] of bounds) {
    const value: unknown = addr[key];
    if (!(Number.isInteger(value) && (value as number) >= 1 && (value as number) <= max)) {
      throw new ExcelInvalidAddressError(
        `{ row: ${String(addr.row)}, col: ${String(addr.col)} }`,
        `a cell position's ${key} is a whole number from 1 to ${max}`,
        { field: `addr.${key}`, where }
      );
    }
  }
}

/**
 * Refuse something that is not a sheet. A streaming worksheet writer is one: the core resolvers ask it for its rows
 * (`getRow` / `findRow`), so it serves every function here through the same code — see {@link editAt} and
 * {@link readAt} for the rows it has already written.
 */
function checkSheet(ws: WorksheetData | StreamWorksheet): asserts ws is WorksheetData {
  if ((ws as { _rows?: unknown } | null | undefined)?._rows === undefined) {
    checkWorksheet(ws);
  }
}

/** The named-style lookup a streaming workbook writer carries (`Workbook.getCellStyle` reads it). */
interface StreamWorkbookStyles {
  _getCellStyle(name: string): { name: string } | undefined;
}

/** Refuse an edit of the cell at `(r, c)` once a streaming writer has written its row. */
function editAt(ws: WorksheetData, where: string, r: string | number, c: number | undefined): void {
  if (asStreamWorksheet(ws) !== undefined) {
    checkStreamRowsOpen(ws, where, resolveAddress(ws, r, c).row);
  }
}

/** Refuse a read of the cell at `(r, c)` once a streaming writer has written its row: it no longer holds it. */
function readAt(ws: WorksheetData, where: string, r: string | number, c: number | undefined): void {
  if (asStreamWorksheet(ws) !== undefined) {
    checkStreamRowsHeld(ws, where, resolveAddress(ws, r, c).row);
  }
}

/** The cell an address names, created when it is absent — for {@link getStyle}, which hands out its live style. */
function target(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  addr: string | CellPosition | number,
  col: number | undefined
): CellData {
  checkSheet(ws);
  const [r, c] = ref(where, addr, col, col);
  editAt(ws, where, r, c);
  return getCell(ws, r, c);
}

/**
 * {@link target} for a reader that answers with a value: fails on the same malformed references, but never creates
 * the row or the cell — `undefined` when nothing is there, and the reader answers as for an empty cell.
 *
 * A read must not change what the workbook saves — materialising widened `rowCount` and `<dimension>`,
 * and in a styled row or column wrote a styled empty cell into the file. {@link getStyle} is the one
 * deliberate exception: it hands back the live style record for the caller to edit in place
 * (`Cell.getStyle(ws, "B2").font = …`), and there is nothing to edit until the cell exists.
 */
function peek(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  addr: string | CellPosition | number,
  col: number | undefined
): CellData | undefined {
  checkSheet(ws);
  const [r, c] = ref(where, addr, col, col);
  readAt(ws, where, r, c);
  return locateCell(ws, r, c);
}

/** The address arguments resolved to an address, for a reader that answers from the address alone. */
function addressOf(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  addr: string | CellPosition | number,
  col: number | undefined
): DecodedAddress {
  checkSheet(ws);
  const [r, c] = ref(where, addr, col, col);
  return resolveAddress(ws, r, c);
}

/**
 * A copy of one style facet of a cell, for the per-facet readers. The cell is previewed, never created, and the
 * answer is a copy whether or not the cell exists — so editing it changes nothing, as with `Row`/`Column`. A facet is
 * changed with its setter, or in place through the live record {@link getStyle} returns.
 */
function facet<K extends keyof Style>(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  addr: string | CellPosition | number,
  col: number | undefined,
  key: K
): Style[K] | undefined {
  checkSheet(ws);
  const [r, c] = ref(where, addr, col, col);
  readAt(ws, where, r, c);
  const value = previewCell(ws, r, c).style[key];
  return value === undefined ? undefined : copyStyleFacet(key, value);
}

/**
 * The address and the value of a setter's `(addr, value)` or deprecated `(row, col, value)` arguments, the form told
 * apart by {@link ref}.
 */
function addressWithValue<V>(
  where: string,
  addr: string | CellPosition | number,
  valueOrCol: V | number,
  value: V | undefined
): [string | number, number | undefined, V] {
  if (typeof addr === "number") {
    return [addr, valueOrCol as number, value as V];
  }
  const [r, c] = ref(where, addr, undefined, value);
  return [r, c, valueOrCol as V];
}

/**
 * The cell and the value of a setter's arguments. `check` sees the value before the cell is looked up, so a refused
 * value does not leave a new empty cell behind.
 */
function targetWithValue<V>(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  addr: string | CellPosition | number,
  valueOrCol: V | number,
  value: V | undefined,
  check?: (value: V, r: string | number, c: number | undefined) => void
): [CellData, V] {
  checkSheet(ws);
  const [r, c, resolved] = addressWithValue(where, addr, valueOrCol, value);
  check?.(resolved, r, c);
  editAt(ws, where, r, c);
  return [getCell(ws, r, c), resolved];
}

/**
 * A name a cell is given is a defined name, so it is checked as `DefinedNames.add` checks one — before the cell is
 * looked up, so a refused name leaves nothing behind.
 */
function checkCellName(ws: WorksheetData | StreamWorksheet, where: string, name: unknown): void {
  definedNamesCheckName(getSheetWorkbook(ws as WorksheetData)._definedNames, where, name);
}

// --- value / type / text ---

/** Read a cell value by `"A1"` address or `{ row, col }`. */
export function getValue(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): CellValue;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getValue(ws, { row, col })`. Removed in 1.0. */
export function getValue(ws: WorksheetData | StreamWorksheet, row: number, col: number): CellValue;
export function getValue(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): CellValue {
  const cell = peek(ws, "Cell.getValue", addr, col);
  return cell ? cellReadValue(cell) : null;
}
/**
 * Set a cell value by `"A1"` address or `{ row, col }`.
 *
 * A value of a known shape is checked as that shape, and a malformed one is refused with `ExcelArgumentError`: rich
 * text (`{ richText: [{ text, font? }] }`), a formula or shared formula (non-empty text; a shared formula's master is a
 * formula cell above or to the left), its `result` (a scalar, a `Date` or `{ error }`), a hyperlink (`hyperlink`,
 * `text` and `tooltip` are text), an error (`{ error: "#N/A" }` — one of Excel's error literals) and a checkbox. An
 * object or array of none of those shapes is stored as its JSON text, as it always has been.
 */
export function setValue(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: CellValueInput
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setValue(ws, { row, col }, …)`. Removed in 1.0. */
export function setValue(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: CellValueInput
): void;
export function setValue(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: CellValueInput,
  value?: CellValueInput
): void {
  const [cell, resolved] = targetWithValue<CellValueInput>(
    ws,
    "Cell.setValue",
    addr,
    valueOrCol,
    value,
    (v, r, c) => {
      const site = { where: "Cell.setValue", field: "value" };
      checkCellValue(v, site);
      // A shared formula's master is checked before the cell is created, so a refusal leaves nothing behind.
      const at = resolveAddress(ws as WorksheetData, r, c);
      checkSharedFormulaAt(ws as WorksheetData, at.row, at.col, v, site);
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
export function getDateParts(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): ExcelDateTimeParts | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getDateParts(ws, { row, col })`. Removed in 1.0. */
export function getDateParts(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): ExcelDateTimeParts | undefined;
export function getDateParts(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): ExcelDateTimeParts | undefined {
  const cell = peek(ws, "Cell.getDateParts", addr, col);
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
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  parts: Partial<ExcelDateTimeParts>,
  kind?: TemporalKind
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setDateParts(ws, { row, col }, …)`. Removed in 1.0. */
export function setDateParts(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  parts: Partial<ExcelDateTimeParts>,
  kind?: TemporalKind
): void;
export function setDateParts(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  partsOrCol: Partial<ExcelDateTimeParts> | number,
  partsOrKind?: Partial<ExcelDateTimeParts> | TemporalKind,
  kind?: TemporalKind
): void {
  checkSheet(ws);
  const legacy = typeof addr === "number";
  const [r, c] = legacy
    ? [addr, partsOrCol as number]
    : ref(
        "Cell.setDateParts",
        addr,
        undefined,
        typeof partsOrCol === "number" ? partsOrCol : undefined
      );
  const parts = (legacy ? partsOrKind : partsOrCol) as Partial<ExcelDateTimeParts>;
  const resolvedKind = (legacy ? kind : partsOrKind) as TemporalKind | undefined;
  // Checked before the cell is looked up, so a refused value leaves no new empty cell behind.
  const check = argumentChecks("Cell.setDateParts");
  check.required("parts", parts);
  check.record("parts", parts);
  check.oneOf("kind", resolvedKind, ["date", "time", "dateTime"]);
  validateDateParts(parts, getSheetWorkbook(ws).properties.date1904 === true);
  editAt(ws, "Cell.setDateParts", r, c);
  const cell = getCell(ws, r, c);
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
export function getDateKind(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): DateFormatKind | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getDateKind(ws, { row, col })`. Removed in 1.0. */
export function getDateKind(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): DateFormatKind | undefined;
export function getDateKind(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): DateFormatKind | undefined {
  const cell = peek(ws, "Cell.getDateKind", addr, col);
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
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  kind?: TemporalKind
): TemporalPlainValue | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getTemporal(ws, { row, col })`. Removed in 1.0. */
export function getTemporal(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  kind?: TemporalKind
): TemporalPlainValue | undefined;
export function getTemporal(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  colOrKind?: number | TemporalKind,
  kind?: TemporalKind
): TemporalPlainValue | undefined {
  const legacy = typeof addr === "number";
  const cell = peek(
    ws,
    "Cell.getTemporal",
    addr,
    typeof colOrKind === "number" ? colOrKind : undefined
  );
  return cell ? cellGetTemporal(cell, (legacy ? kind : colOrKind) as TemporalKind) : undefined;
}

/** Read a cell's text by `"A1"` address or `{ row, col }`. */
export function getText(ws: WorksheetData | StreamWorksheet, addr: string | CellPosition): string;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getText(ws, { row, col })`. Removed in 1.0. */
export function getText(ws: WorksheetData | StreamWorksheet, row: number, col: number): string;
export function getText(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string {
  const cell = peek(ws, "Cell.getText", addr, col);
  return cell ? cellText(cell) : "";
}
/** Read a cell's display text by `"A1"` address or `{ row, col }`. */
export function getDisplayText(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getDisplayText(ws, { row, col })`. Removed in 1.0. */
export function getDisplayText(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): string;
export function getDisplayText(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string {
  const cell = peek(ws, "Cell.getDisplayText", addr, col);
  return cell ? cellDisplayText(cell) : "";
}
/** Read a cell's value kind by `"A1"` address or `{ row, col }`. */
export function getType(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): ValueType;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getType(ws, { row, col })`. Removed in 1.0. */
export function getType(ws: WorksheetData | StreamWorksheet, row: number, col: number): ValueType;
export function getType(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): ValueType {
  const cell = peek(ws, "Cell.getType", addr, col);
  return cell ? cellType(cell) : ValueType.Null;
}
/** Read a cell's effective value kind by `"A1"` address or `{ row, col }`. */
export function getEffectiveType(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): ValueType;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getEffectiveType(ws, { row, col })`. Removed in 1.0. */
export function getEffectiveType(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): ValueType;
export function getEffectiveType(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): ValueType {
  const cell = peek(ws, "Cell.getEffectiveType", addr, col);
  return cell ? cellEffectiveType(cell) : ValueType.Null;
}

// --- formula ---

/**
 * Read a cell's formula by `"A1"` address or `{ row, col }`.
 *
 * The text is what the cell holds. For a formula read from a file that is the stored form, with the prefixes the file
 * carries (`_xlfn.XLOOKUP(…)`, `_xlpm.x`, `_xleta.SUM`) — including for a formula that was set bare and then written.
 */
export function getFormula(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getFormula(ws, { row, col })`. Removed in 1.0. */
export function getFormula(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): string | undefined;
export function getFormula(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string | undefined {
  const cell = peek(ws, "Cell.getFormula", addr, col);
  return cell ? cellFormula(cell) : undefined;
}
/** Read a formula cell's cached result by `"A1"` address or `{ row, col }`. */
export function getResult(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): FormulaResult | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getResult(ws, { row, col })`. Removed in 1.0. */
export function getResult(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): FormulaResult | undefined;
export function getResult(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): FormulaResult | undefined {
  const cell = peek(ws, "Cell.getResult", addr, col);
  return cell ? ownValue(cellResult(cell)) : undefined;
}

// --- style ---

/**
 * The live style of the cell at `"A1"`, for editing in place — it creates the cell when there is none.
 * To read without creating anything, use {@link findStyle}.
 */
export function getStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Partial<Style>;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getStyle(ws, { row, col })`. Removed in 1.0. */
export function getStyle(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Partial<Style>;
export function getStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Partial<Style> {
  return cellOwnStyle(target(ws, "Cell.getStyle", addr, col));
}
/**
 * A copy of the style of the cell at `"A1"`, or `undefined` when the sheet has no cell there — the
 * read-only counterpart of {@link getStyle}, which hands out the live style for editing and so creates
 * the cell. This creates nothing, and editing what it returns changes nothing.
 */
export function findStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Partial<Style> | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.findStyle(ws, { row, col })`. Removed in 1.0. */
export function findStyle(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Partial<Style> | undefined;
export function findStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Partial<Style> | undefined {
  const cell = peek(ws, "Cell.findStyle", addr, col);
  return cell ? (copyStyle(cellGetStyle(cell)) as Partial<Style>) : undefined;
}
/** Merge a partial style into the cell at "A1" address. */
export function setStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  style: Partial<Style>
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setStyle(ws, { row, col }, …)`. Removed in 1.0. */
export function setStyle(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  style: Partial<Style>
): void;
export function setStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  styleOrCol: Partial<Style> | number,
  style?: Partial<Style>
): void {
  const [cell, resolved] = targetWithValue<Partial<Style>>(
    ws,
    "Cell.setStyle",
    addr,
    styleOrCol,
    style,
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
export function applyStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  name: string
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.applyStyle(ws, { row, col }, …)`. Removed in 1.0. */
export function applyStyle(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  name: string
): void;
export function applyStyle(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  nameOrCol: string | number,
  name?: string
): void {
  checkSheet(ws);
  const [r, c, styleName] = addressWithValue<string>("Cell.applyStyle", addr, nameOrCol, name);
  argumentChecks("Cell.applyStyle").required("name", styleName);
  argumentChecks("Cell.applyStyle").string("name", styleName, { nonEmpty: true });
  // Found in any case, as Excel finds it; the cell records the style's own spelling. A streaming writer keeps its
  // styles itself, where `Workbook.getCellStyle` reads them for it.
  const writer = asStreamWorksheet(ws);
  const named = writer
    ? (writer.workbook as unknown as StreamWorkbookStyles)._getCellStyle(styleName)
    : getCellStyle(getSheetWorkbook(ws), styleName);
  if (!named) {
    throw new ExcelArgumentError(
      "name",
      `Named cell style "${styleName}" is not defined. Define it first with Workbook.addCellStyle().`
    );
  }
  editAt(ws, "Cell.applyStyle", r, c);
  cellSetStyle(getCell(ws, r, c), { styleName: named.name });
}

/** @deprecated Use {@link applyStyle}. Removed in 1.0. */
export const applyCellStyle = applyStyle;

// --- merge ---

/** Test whether a cell is merged by `"A1"` address or `{ row, col }`. */
export function isMerged(ws: WorksheetData | StreamWorksheet, addr: string | CellPosition): boolean;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.isMerged(ws, { row, col })`. Removed in 1.0. */
export function isMerged(ws: WorksheetData | StreamWorksheet, row: number, col: number): boolean;
export function isMerged(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): boolean {
  const cell = peek(ws, "Cell.isMerged", addr, col);
  return cell ? cellIsMerged(cell) : false;
}
/** The address of the merge's master (top-left) cell, or the cell's own address when it is not merged. */
export function getMergeMasterAddress(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getMergeMasterAddress(ws, { row, col })`. Removed in 1.0. */
export function getMergeMasterAddress(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): string;
export function getMergeMasterAddress(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string {
  const at = addressOf(ws, "Cell.getMergeMasterAddress", addr, col);
  readAt(ws as WorksheetData, "Cell.getMergeMasterAddress", at.row, at.col);
  const cell = locateCell(ws as WorksheetData, at.address);
  return cell ? cellMaster(cell).address : at.address;
}

/** @deprecated Use {@link getMergeMasterAddress}. Removed in 1.0. */
export function getMergeMaster(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): CellData;
/** @deprecated Use {@link getMergeMasterAddress}. Removed in 1.0. */
export function getMergeMaster(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): CellData;
export function getMergeMaster(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): CellData {
  checkSheet(ws);
  // A merged cell always exists, so an absent one is its own master: preview it rather than create it.
  const [r, c] = ref("Cell.getMergeMaster", addr, col, col);
  readAt(ws, "Cell.getMergeMaster", r, c);
  return cellMaster(previewCell(ws, r, c));
}

// --- hyperlink ---

/** Read a cell's hyperlink by `"A1"` address or `{ row, col }`. */
export function getHyperlink(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getHyperlink(ws, { row, col })`. Removed in 1.0. */
export function getHyperlink(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): string | undefined;
export function getHyperlink(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string | undefined {
  const cell = peek(ws, "Cell.getHyperlink", addr, col);
  return cell ? cellHyperlink(cell) : undefined;
}

// --- note ---

/**
 * Read a cell's note by `"A1"` address or `{ row, col }`: its text, or the note configuration it was given — the
 * author is {@link getNoteAuthor}. `undefined` when the cell has no note. Creates nothing.
 *
 * Excel calls these notes; a threaded comment is a different, newer object this API does not write.
 */
export function getNote(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string | NoteConfig | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getNote(ws, { row, col })`. Removed in 1.0. */
export function getNote(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): string | NoteConfig | undefined;
export function getNote(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string | NoteConfig | undefined {
  const cell = peek(ws, "Cell.getNote", addr, col);
  return cell ? ownValue(cellNote(cell)) : undefined;
}

/** The author of a cell's note, or `undefined` when the note has none or the cell has no note. Creates nothing. */
export function getNoteAuthor(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string | undefined {
  const cell = peek(ws, "Cell.getNoteAuthor", addr, undefined);
  return cell ? (cellComment(cell)?.author ?? undefined) : undefined;
}

/** Options for {@link setNote}. */
export interface SetNoteOptions {
  /** The note's author, shown by Excel above its text. Omitted, the note has none. */
  readonly author?: string;
}

/**
 * Set a cell's note by `"A1"` address or `{ row, col }`, replacing any it had — its text or a note configuration,
 * and in `options` its author. Remove one with {@link removeNote}.
 */
export function setNote(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  note: string | NoteConfig,
  options?: SetNoteOptions
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setNote(ws, { row, col }, …)`. Removed in 1.0. */
export function setNote(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  note: string | NoteConfig
): void;
export function setNote(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  noteOrCol: string | NoteConfig | number,
  noteOrOptions?: string | NoteConfig | SetNoteOptions
): void {
  const where = "Cell.setNote";
  const positional = typeof addr === "number";
  const options = positional ? undefined : (noteOrOptions as SetNoteOptions | undefined);
  if (!positional) {
    const check = argumentChecks(where);
    if (check.record("options", options)) {
      check.string("options.author", options.author);
    }
  }
  const [cell, resolved] = targetWithValue<string | NoteConfig>(
    ws,
    where,
    addr,
    noteOrCol,
    positional ? (noteOrOptions as string | NoteConfig | undefined) : undefined,
    value => {
      if (typeof value !== "string" && (typeof value !== "object" || value === null)) {
        argumentChecks(where).fail("note", "a string or a note configuration", value);
      }
      checkNote(where, "note", value);
    }
  );
  cellSetNote(cell, resolved, options?.author);
}

/**
 * Remove a cell's note by `"A1"` address or `{ row, col }`. Returns whether there was one; creates nothing.
 */
export function removeNote(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): boolean {
  const where = "Cell.removeNote";
  checkSheet(ws);
  const [r, c] = ref(where, addr, undefined, undefined);
  editAt(ws, where, r, c);
  const cell = locateCell(ws, r, c);
  if (!cell || cellComment(cell) === undefined) {
    return false;
  }
  cellSetComment(cell, undefined);
  return true;
}

// --- defined names ---

/** Read the defined names covering a cell by `"A1"` address or `{ row, col }`. */
export function getNames(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string[];
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getNames(ws, { row, col })`. Removed in 1.0. */
export function getNames(ws: WorksheetData | StreamWorksheet, row: number, col: number): string[];
export function getNames(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string[] {
  // Defined names live on the workbook, keyed by address; the cell itself holds none.
  const address = addressOf(ws, "Cell.getNames", addr, col);
  return definedNamesGetNamesEx(getSheetWorkbook(ws as WorksheetData)._definedNames, address);
}
/** Add a defined name to a cell by `"A1"` address or `{ row, col }`. */
export function addName(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  name: string
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.addName(ws, { row, col }, …)`. Removed in 1.0. */
export function addName(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  name: string
): void;
export function addName(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  nameOrCol: string | number,
  name?: string
): void {
  const [cell, resolved] = targetWithValue<string>(
    ws,
    "Cell.addName",
    addr,
    nameOrCol,
    name,
    value => checkCellName(ws, "Cell.addName", value)
  );
  cellAddName(cell, resolved);
}
/**
 * Remove a defined name from the cell at `"A1"`. Returns `true` if the cell carried it. Creates nothing: the
 * names live on the workbook, keyed by address.
 */
export function removeName(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  name: string
): boolean;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.removeName(ws, { row, col }, …)`. Removed in 1.0. */
export function removeName(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  name: string
): boolean;
export function removeName(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  nameOrCol: string | number,
  name?: string
): boolean {
  // The names live on the workbook, keyed by sheet name: through a removed sheet's handle this would edit the names
  // of whichever sheet has taken the name since.
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  const [r, c, resolved] = addressWithValue<string>("Cell.removeName", addr, nameOrCol, name);
  const address = resolveAddress(ws, r, c);
  argumentChecks("Cell.removeName").string("name", resolved);
  return definedNamesRemoveEx(getSheetWorkbook(ws)._definedNames, address, resolved);
}
/**
 * Replace a cell's defined names with one by `"A1"` address or `{ row, col }`.
 *
 * @deprecated Use {@link setNames} with a one-name array: `setNames(ws, addr, [name])`. Removed in 1.0.
 */
export function setName(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  name: string
): void;
/**
 * Replace a cell's defined names with one by 1-based (row, col).
 *
 * @deprecated Use {@link setNames} with a one-name array: `setNames(ws, row, col, [name])`. Removed in 1.0.
 */
export function setName(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  name: string
): void;
export function setName(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  nameOrCol: string | number,
  name?: string
): void {
  const [cell, resolved] = targetWithValue<string>(
    ws,
    "Cell.setName",
    addr,
    nameOrCol,
    name,
    value => checkCellName(ws, "Cell.setName", value)
  );
  cellSetName(cell, resolved);
}
/** Replace a cell's defined names by `"A1"` address or `{ row, col }`. */
export function setNames(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  names: readonly string[]
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setNames(ws, { row, col }, …)`. Removed in 1.0. */
export function setNames(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  names: readonly string[]
): void;
export function setNames(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  namesOrCol: readonly string[] | number,
  names?: readonly string[]
): void {
  const [cell, resolved] = targetWithValue<readonly string[]>(
    ws,
    "Cell.setNames",
    addr,
    namesOrCol,
    names,
    value => {
      if (!Array.isArray(value)) {
        argumentChecks("Cell.setNames").fail("names", "an array of names", value);
      }
      value.forEach(n => checkCellName(ws, "Cell.setNames", n));
    }
  );
  cellSetNames(cell, [...resolved]);
}
/** Remove every defined name from a cell by `"A1"` address or `{ row, col }`. */
export function clearNames(ws: WorksheetData | StreamWorksheet, addr: string | CellPosition): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.clearNames(ws, { row, col }, …)`. Removed in 1.0. */
export function clearNames(ws: WorksheetData | StreamWorksheet, row: number, col: number): void;
export function clearNames(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): void {
  // See `removeName`: the names are keyed by sheet name, so a removed sheet's handle is refused.
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  definedNamesRemoveAllNames(
    getSheetWorkbook(ws)._definedNames,
    addressOf(ws, "Cell.clearNames", addr, col)
  );
}
/** @deprecated Use {@link clearNames}. Removed in 1.0. */
export const removeAllNames = clearNames;

// --- data validation ---

/** A copy of the validation rule covering a cell, by `"A1"` address or `{ row, col }`, or `undefined`. */
export function getValidation(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): DataValidationRule | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getValidation(ws, { row, col })`. Removed in 1.0. */
export function getValidation(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): DataValidationRule | undefined;
export function getValidation(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): DataValidationRule | undefined {
  // Validation rules live on the sheet, keyed by address; the cell itself holds none. A copy, as the other facet
  // readers return: editing it changes nothing until it is set again.
  const address = addressOf(ws, "Cell.getValidation", addr, col);
  const rule = dataValidationFind(ws.dataValidations, address.address);
  return rule === undefined ? undefined : structuredClone(rule);
}
/** Set a cell's validation rule by `"A1"` address or `{ row, col }`. The sheet keeps a copy of `value`. */
export function setValidation(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: DataValidationRule
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setValidation(ws, { row, col }, …)`. Removed in 1.0. */
export function setValidation(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: DataValidationRule
): void;
export function setValidation(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: DataValidationRule | number,
  value?: DataValidationRule
): void {
  const [cell, resolved] = targetWithValue<DataValidationRule>(
    ws,
    "Cell.setValidation",
    addr,
    valueOrCol,
    value,
    rule => checkDataValidationRule("Cell.setValidation", rule)
  );
  // A copy: the caller's object used to be the stored rule, so one rule given to two cells — or edited afterwards —
  // changed what every one of them wrote.
  cellSetDataValidation(cell, resolved === undefined ? resolved : structuredClone(resolved));
}

// --- model (advanced / round-trip) ---

/**
 * A copy of a cell's round-trip model by `"A1"` address or `{ row, col }` — the shape {@link setModel} applies. Editing
 * it changes nothing; apply a changed copy with {@link setModel}, or edit the cell through the setters. Creates
 * nothing: an absent cell reports what it would inherit.
 */
export function getModel(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): CellModel;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getModel(ws, { row, col })`. Removed in 1.0. */
export function getModel(ws: WorksheetData | StreamWorksheet, row: number, col: number): CellModel;
export function getModel(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): CellModel {
  const where = "Cell.getModel";
  checkSheet(ws);
  const [r, c] = ref(where, addr, col, col);
  readAt(ws, where, r, c);
  return copyPlain(cellGetModel(previewCell(ws, r, c)));
}
/**
 * Apply a round-trip cell model by `"A1"` address or `{ row, col }`.
 *
 * @throws {ExcelArgumentError} when the model is for another address, has an unknown `type`, or holds a value of
 * another kind than its `type` or one no workbook can hold (`NaN`, an invalid `Date`).
 */
export function setModel(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  model: CellModel
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setModel(ws, { row, col }, …)`. Removed in 1.0. */
export function setModel(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  model: CellModel
): void;
export function setModel(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  modelOrCol: CellModel | number,
  model?: CellModel
): void {
  // The address is resolved first, so a model for another cell is refused before this one is created: it was
  // written into this cell's `<row>` under its own address, and the file could not be read back.
  const [cell, resolved] = targetWithValue<CellModel>(
    ws,
    "Cell.setModel",
    addr,
    modelOrCol,
    model,
    (value, r, c) =>
      checkCellModel(
        "Cell.setModel",
        "model",
        value,
        resolveAddress(ws as WorksheetData, r, c).address
      )
  );
  cellSetModel(cell, copyPlain(resolved));
}

// --- individual style facets (getters + setters) ---

/** A copy of a cell's font by `"A1"` address or `{ row, col }`. Creates nothing; editing the copy changes nothing — change it with {@link setFont}. */
export function getFont(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Partial<Font> | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getFont(ws, { row, col })`. Removed in 1.0. */
export function getFont(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Partial<Font> | undefined;
export function getFont(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Partial<Font> | undefined {
  return facet(ws, "Cell.getFont", addr, col, "font");
}
/** Set a cell's font by `"A1"` address or `{ row, col }`. */
export function setFont(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: Partial<Font> | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setFont(ws, { row, col }, …)`. Removed in 1.0. */
export function setFont(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: Partial<Font> | undefined
): void;
export function setFont(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: Partial<Font> | undefined | number,
  value?: Partial<Font>
): void {
  const [cell, resolved] = targetWithValue<Partial<Font> | undefined>(
    ws,
    "Cell.setFont",
    addr,
    valueOrCol,
    value,
    value => checkStyleFacet("Cell.setFont", "font", value, "value")
  );
  cellSetFont(cell, resolved);
}
/** Read a cell's number format by `"A1"` address or `{ row, col }`. */
export function getNumFmt(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): string | NumFmt | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getNumFmt(ws, { row, col })`. Removed in 1.0. */
export function getNumFmt(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): string | NumFmt | undefined;
export function getNumFmt(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): string | NumFmt | undefined {
  const cell = peek(ws, "Cell.getNumFmt", addr, col);
  if (cell) {
    const numFmt = cellNumFmt(cell);
    return numFmt === undefined ? undefined : copyStyleFacet("numFmt", numFmt);
  }
  const { row, col: column } = addressOf(ws, "Cell.getNumFmt", addr, col);
  const inherited = inheritedNumFmt(ws as WorksheetData, row, column);
  return inherited === undefined ? undefined : copyStyleFacet("numFmt", inherited);
}
/** Set a cell's number format by `"A1"` address or `{ row, col }`. */
export function setNumFmt(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: string | NumFmt | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setNumFmt(ws, { row, col }, …)`. Removed in 1.0. */
export function setNumFmt(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: string | NumFmt | undefined
): void;
export function setNumFmt(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: string | NumFmt | undefined | number,
  value?: string | NumFmt
): void {
  const [cell, resolved] = targetWithValue<string | NumFmt | undefined>(
    ws,
    "Cell.setNumFmt",
    addr,
    valueOrCol,
    value,
    value => checkStyleFacet("Cell.setNumFmt", "numFmt", value, "value")
  );
  cellSetNumFmt(cell, resolved);
}
/** A copy of a cell's alignment by `"A1"` address or `{ row, col }`. Creates nothing; editing the copy changes nothing — change it with {@link setAlignment}. */
export function getAlignment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Partial<Alignment> | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getAlignment(ws, { row, col })`. Removed in 1.0. */
export function getAlignment(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Partial<Alignment> | undefined;
export function getAlignment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Partial<Alignment> | undefined {
  return facet(ws, "Cell.getAlignment", addr, col, "alignment");
}
/** Set a cell's alignment by `"A1"` address or `{ row, col }`. */
export function setAlignment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: Partial<Alignment> | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setAlignment(ws, { row, col }, …)`. Removed in 1.0. */
export function setAlignment(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: Partial<Alignment> | undefined
): void;
export function setAlignment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: Partial<Alignment> | undefined | number,
  value?: Partial<Alignment>
): void {
  const [cell, resolved] = targetWithValue<Partial<Alignment> | undefined>(
    ws,
    "Cell.setAlignment",
    addr,
    valueOrCol,
    value,
    value => checkStyleFacet("Cell.setAlignment", "alignment", value, "value")
  );
  cellSetAlignment(cell, resolved);
}
/** A copy of a cell's borders by `"A1"` address or `{ row, col }`. Creates nothing; editing the copy changes nothing — change it with {@link setBorder}. */
export function getBorder(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Partial<Borders> | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getBorder(ws, { row, col })`. Removed in 1.0. */
export function getBorder(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Partial<Borders> | undefined;
export function getBorder(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Partial<Borders> | undefined {
  return facet(ws, "Cell.getBorder", addr, col, "border");
}
/** Set a cell's borders by `"A1"` address or `{ row, col }`. */
export function setBorder(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: Partial<Borders> | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setBorder(ws, { row, col }, …)`. Removed in 1.0. */
export function setBorder(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: Partial<Borders> | undefined
): void;
export function setBorder(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: Partial<Borders> | undefined | number,
  value?: Partial<Borders>
): void {
  const [cell, resolved] = targetWithValue<Partial<Borders> | undefined>(
    ws,
    "Cell.setBorder",
    addr,
    valueOrCol,
    value,
    value => checkStyleFacet("Cell.setBorder", "border", value, "value")
  );
  cellSetBorder(cell, resolved);
}
/** A copy of a cell's fill by `"A1"` address or `{ row, col }`. Creates nothing; editing the copy changes nothing — change it with {@link setFill}. */
export function getFill(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Fill | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getFill(ws, { row, col })`. Removed in 1.0. */
export function getFill(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Fill | undefined;
export function getFill(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Fill | undefined {
  return facet(ws, "Cell.getFill", addr, col, "fill");
}
/** Set a cell's fill by `"A1"` address or `{ row, col }`. */
export function setFill(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: Fill | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setFill(ws, { row, col }, …)`. Removed in 1.0. */
export function setFill(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: Fill | undefined
): void;
export function setFill(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: Fill | undefined | number,
  value?: Fill
): void {
  const [cell, resolved] = targetWithValue<Fill | undefined>(
    ws,
    "Cell.setFill",
    addr,
    valueOrCol,
    value,
    value => checkStyleFacet("Cell.setFill", "fill", value, "value")
  );
  cellSetFill(cell, resolved);
}
/** A copy of a cell's protection by `"A1"` address or `{ row, col }`. Creates nothing; editing the copy changes nothing — change it with {@link setProtection}. */
export function getProtection(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): Partial<Protection> | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getProtection(ws, { row, col })`. Removed in 1.0. */
export function getProtection(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): Partial<Protection> | undefined;
export function getProtection(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): Partial<Protection> | undefined {
  return facet(ws, "Cell.getProtection", addr, col, "protection");
}
/** Set a cell's protection by `"A1"` address or `{ row, col }`. */
export function setProtection(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: Partial<Protection> | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setProtection(ws, { row, col }, …)`. Removed in 1.0. */
export function setProtection(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: Partial<Protection> | undefined
): void;
export function setProtection(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: Partial<Protection> | undefined | number,
  value?: Partial<Protection>
): void {
  const [cell, resolved] = targetWithValue<Partial<Protection> | undefined>(
    ws,
    "Cell.setProtection",
    addr,
    valueOrCol,
    value,
    value => checkStyleFacet("Cell.setProtection", "protection", value, "value")
  );
  cellSetProtection(cell, resolved);
}

// --- comment (deprecated: the note with its author) ---

/**
 * @deprecated A comment is the cell's note with its author: read them with {@link getNote} and
 * {@link getNoteAuthor}. Removed in 1.0.
 */
export function getComment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): NoteData | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getComment(ws, { row, col })`. Removed in 1.0. */
export function getComment(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): NoteData | undefined;
export function getComment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): NoteData | undefined {
  const cell = peek(ws, "Cell.getComment", addr, col);
  return cell ? cellComment(cell) : undefined;
}
/**
 * @deprecated A comment is the cell's note with its author: `Cell.setNote(ws, addr, note, { author })`, and
 * {@link removeNote} for `undefined`. Removed in 1.0.
 */
export function setComment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  comment: NoteData | NoteConfig | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setComment(ws, { row, col }, …)`. Removed in 1.0. */
export function setComment(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  comment: NoteData | NoteConfig | undefined
): void;
export function setComment(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  commentOrCol: NoteData | NoteConfig | undefined | number,
  comment?: NoteData | NoteConfig
): void {
  const [cell, resolved] = targetWithValue<NoteData | NoteConfig | undefined>(
    ws,
    "Cell.setComment",
    addr,
    commentOrCol,
    comment,
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
 * Set a formula cell's cached result by `"A1"` address or `{ row, col }` — what a reader shows until the workbook is recalculated.
 *
 * @throws {ExcelArgumentError} when the cell holds no formula (passing `undefined` there is allowed and does nothing).
 */
export function setResult(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition,
  value: FormulaResult | undefined
): void;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.setResult(ws, { row, col }, …)`. Removed in 1.0. */
export function setResult(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number,
  value: FormulaResult | undefined
): void;
export function setResult(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  valueOrCol: FormulaResult | undefined | number,
  value?: FormulaResult
): void {
  const [cell, resolved] = targetWithValue<FormulaResult | undefined>(
    ws,
    "Cell.setResult",
    addr,
    valueOrCol,
    value,
    (result, r, c) => {
      checkFormulaResult(result, { where: "Cell.setResult", field: "value" });
      // A cached result belongs to a formula. On any other cell it was dropped without a word.
      const existing = locateCell(ws as WorksheetData, r, c);
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
/** The cell's fully-qualified address by `"A1"` address or `{ row, col }`. */
export function getFullAddress(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): DecodedAddress;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getFullAddress(ws, { row, col })`. Removed in 1.0. */
export function getFullAddress(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): DecodedAddress;
export function getFullAddress(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): DecodedAddress {
  return addressOf(ws, "Cell.getFullAddress", addr, col);
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
export function has(ws: WorksheetData | StreamWorksheet, addr: string | CellPosition): boolean;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.has(ws, { row, col })`. Removed in 1.0. */
export function has(ws: WorksheetData | StreamWorksheet, row: number, col: number): boolean;
export function has(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): boolean {
  return peek(ws, "Cell.has", addr, col) !== undefined;
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
export function getView(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): CellView | undefined;
/** @deprecated Pass the cell as `{ row, col }`: `Cell.getView(ws, { row, col })`. Removed in 1.0. */
export function getView(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): CellView | undefined;
export function getView(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): CellView | undefined {
  const cell = peek(ws, "Cell.getView", addr, col);
  return cell && cellView(cell);
}

// --- deprecated cell handles ---

/**
 * @deprecated Use {@link has} to test for a cell, {@link getView} to read one
 * without creating it, and address it by position. Removed in 1.0.
 */
export function find(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition
): CellData | undefined;
/** @deprecated Use {@link has} / {@link getView}. Removed in 1.0. */
export function find(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: number
): CellData | undefined;
export function find(
  ws: WorksheetData | StreamWorksheet,
  addr: string | CellPosition | number,
  col?: number
): CellData | undefined {
  return peek(ws, "Cell.find", addr, col);
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

/** @deprecated Use the flat `CellValue` type. Removed in 1.0. */
export type Value = CellValue;
