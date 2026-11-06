/**
 * `Worksheet` namespace surface — sheet-level structure operations.
 *
 * `import { Worksheet } from "documonster/excel"` → `Worksheet.merge(ws, "A1:B2")`,
 * `Worksheet.freeze(ws, { rows: 1 })`, `Worksheet.getDimensions(ws)`, …
 *
 * Row and column work — adding, inserting, splicing, counting, iterating,
 * auto-fitting — lives on `Row` and `Column`; Cell / Chart / Table / Image /
 * Pivot / ConditionalFormatting operations live in their own namespaces too.
 */
import { checkCellValue } from "@excel/core/cell";
import type { RangeData, RangeInput } from "@excel/core/range";
import { rangeCreate } from "@excel/core/range";
import { rowGetValues } from "@excel/core/row";
import { refuseLongSheetName, validateSheetName } from "@excel/core/workbook-core";
import {
  toJSON as coreToJson,
  addAOA,
  addJSON,
  fillFormula as coreFillFormula,
  freeze as coreFreeze,
  protect as coreProtect,
  unprotect as coreUnprotect,
  setAutoFilter as coreSetAutoFilter,
  split as coreSplit,
  unMergeCells,
  autoFitColumn as coreAutoFitColumn,
  autoFitColumns as coreAutoFitColumns,
  autoFitRow as coreAutoFitRow,
  autoFitRows as coreAutoFitRows,
  destroy as coreDestroy,
  duplicateRow as coreDuplicateRow,
  getActualColumnCount as coreGetActualColumnCount,
  getActualRowCount as coreGetActualRowCount,
  getColumnCount as coreGetColumnCount,
  getColumnDefinitions as coreGetColumnDefinitions,
  getColumns as coreGetColumns,
  getLastColumn as coreGetLastColumn,
  getLastRow as coreGetLastRow,
  getMergedRegions as coreGetMergedRegions,
  getRowCount as coreGetRowCount,
  getSheetDimensions as coreGetSheetDimensions,
  insertRow as coreInsertRow,
  insertRows as coreInsertRows,
  panes as corePanes,
  removeConditionalFormatting as coreRemoveConditionalFormatting,
  setColumns as coreSetColumns,
  spliceColumns as coreSpliceColumns,
  spliceRows as coreSpliceRows,
  getSheetModel,
  mergeCells,
  mergeCellsWithoutStyle,
  setSheetName
} from "@excel/core/worksheet";
import type {
  AddAoaOptions,
  AddJsonOptions,
  FillFormulaOptions,
  ProtectSheetOptions,
  SheetToJsonOptions,
  WorksheetModel
} from "@excel/core/worksheet";
import {
  addRow as coreAddRow,
  addRows as coreAddRows,
  autoFilter as coreAutoFilter,
  eachRow as coreEachRow,
  findRow as coreFindRow,
  findRows as coreFindRows,
  checkWorksheet,
  checkWorksheetAttached,
  getRow as coreGetRow,
  getRows as coreGetRows,
  getSheetName,
  getSheetValues,
  getSheetWorkbook
} from "@excel/core/worksheet-core";
import type {
  FreezePanesOptions,
  SheetProtection,
  SplitPanesOptions,
  WorksheetData,
  WorksheetPanes
} from "@excel/core/worksheet-core";
import { setSheetModel } from "@excel/core/worksheet-load";
import {
  ExcelColumnOutOfBoundsError,
  ExcelInvalidAddressError,
  ExcelRowOutOfBoundsError
} from "@excel/errors";
import { checkAutoFilter } from "@excel/surface/argument-checks";
import { add as addConditionalFormattingChecked } from "@excel/surface/conditional-formatting";
import { checkWorksheetModel } from "@excel/surface/model-checks";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import { asStreamWorksheet, refuseStreamWorksheet } from "@excel/surface/stream-targets";
import type { AutoFilter, CellValue, CellValueInput } from "@excel/types";
import { decodeCell } from "@excel/utils/address";
import type { Origin } from "@excel/utils/address";
import type { ArgumentChecks } from "@excel/utils/arguments";
import {
  argumentChecks,
  checkProtectCall,
  checkSchema,
  normalizeReference
} from "@excel/utils/arguments";
import { copyPlain } from "@excel/utils/copy-plain";
import type { ValuesOptions } from "@excel/utils/numbered-values";
import { splitPasswordOptions } from "@excel/utils/sheet-protection";
import { sheetNameKey } from "@utils/name-key";
import { isRecordArgument } from "@utils/object";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

export {
  verifyPassword,
  toAOA as toAoa,
  getHasMerges as hasMerges,
  setSheetName as setName,
  unfreeze
} from "@excel/core/worksheet";

const MAX_ROW = 1048576;
const MAX_COL = 16384;

// The options of the functions below. A reference (`range`, `origin`) is checked by hand after the schema, because a
// malformed one is an `ExcelInvalidAddressError` rather than an argument error.
const toJsonOptionsSchema = s.object({
  header: s
    .union([s.literal(1, "A"), s.array(s.string())], { expected: '1, "A" or an array of names' })
    .optional(),
  range: s.unknown<number | string>().optional(),
  raw: s.boolean().optional(),
  defaultValue: s.unknown<CellValue>().optional(),
  blankRows: s.boolean().optional(),
  dateFormat: s.string().optional()
});
const _toJsonOptionsMatch: SchemaMatches<typeof toJsonOptionsSchema, SheetToJsonOptions> = true;
void _toJsonOptionsMatch;

const fillFormulaOptionsSchema = s.object(
  {
    results: s.unknown<NonNullable<FillFormulaOptions["results"]>>().optional(),
    shareType: s.literal("shared", "array").optional()
  },
  { readonly: true }
);
const _fillFormulaOptionsMatch: SchemaMatches<typeof fillFormulaOptionsSchema, FillFormulaOptions> =
  true;
void _fillFormulaOptionsMatch;

const origin = () => s.unknown<Origin>().optional();
const addJsonOptionsSchema = s.object({
  header: s.array(s.string()).optional(),
  skipHeader: s.boolean().optional(),
  origin: origin(),
  nullError: s.boolean().optional()
});
const _addJsonOptionsMatch: SchemaMatches<typeof addJsonOptionsSchema, AddJsonOptions> = true;
void _addJsonOptionsMatch;

const addAoaOptionsSchema = s.object({ origin: origin() });
const _addAoaOptionsMatch: SchemaMatches<typeof addAoaOptionsSchema, AddAoaOptions> = true;
void _addAoaOptionsMatch;

const freezeOptionsSchema = s.object({
  state: s.literal("frozen").optional(),
  columns: s.number({ integer: true, min: 0, max: MAX_COL - 1 }).optional(),
  rows: s.number({ integer: true, min: 0, max: MAX_ROW - 1 }).optional()
});
const _freezeOptionsMatch: SchemaMatches<typeof freezeOptionsSchema, FreezePanesOptions> = true;
void _freezeOptionsMatch;

const splitOptionsSchema = s.object({
  state: s.literal("split").optional(),
  x: s.number({ finite: true, min: 0 }).optional(),
  y: s.number({ finite: true, min: 0 }).optional()
});
const _splitOptionsMatch: SchemaMatches<typeof splitOptionsSchema, SplitPanesOptions> = true;
void _splitOptionsMatch;

/** The sheet's name. */
export function getName(ws: WorksheetData): string {
  checkWorksheet(ws);
  return getSheetName(ws);
}

/** The sheet as an array of rows, each an array of values (`{ header: 1 }`). */
export function toJson(
  ws: WorksheetData,
  options: SheetToJsonOptions & { header: 1 }
): CellValue[][];
/**
 * The sheet as a list of records keyed by its first row — or by `header`'s names, or by column letter with
 * `{ header: "A" }`. `range` limits the read to an `A1:C9` range, or starts it at a 0-based row.
 *
 * @throws {ExcelArgumentError} when `header` is not one of those forms or `range` is a negative or fractional row.
 * @throws {ExcelInvalidAddressError} when `range` is text that is not a bounded range.
 */
export function toJson(
  ws: WorksheetData,
  options?: SheetToJsonOptions
): Record<string, CellValue>[];
export function toJson(
  ws: WorksheetData,
  options?: SheetToJsonOptions
): CellValue[][] | Record<string, CellValue>[] {
  const check = argumentChecks("Worksheet.toJson");
  checkSchema("Worksheet.toJson", "options", toJsonOptionsSchema.optional(), options);
  if (options) {
    const { range } = options;
    if (typeof range === "number") {
      check.number("options.range", range, { min: 0, max: MAX_ROW - 1, integer: true });
    } else if (range !== undefined) {
      check.reference("options.range", range, { range: true });
    }
  }
  return coreToJson(ws, options as SheetToJsonOptions);
}

/** An auto-filter with its references' column letters in capitals; see `normalizeReference`. */
function normalizeAutoFilter(range: AutoFilter | null): AutoFilter | null {
  if (typeof range === "string") {
    return normalizeReference(range);
  }
  if (range === null || typeof range !== "object") {
    return range;
  }
  const corner = <T>(value: T): T =>
    (typeof value === "string" ? normalizeReference(value) : value) as T;
  return { ...range, from: corner(range.from), to: corner(range.to) } as AutoFilter;
}

/**
 * The range a merge or unmerge names, checked: whole cells inside the grid, on this sheet. A `Sheet!` qualifier
 * naming another sheet used to be ignored, so `merge(ws, "Other!C1:D2")` merged C1:D2 here.
 */
function mergeRange(
  where: string,
  ws: WorksheetData | StreamWorksheet,
  cells: RangeInput[]
): RangeData {
  checkWorksheet(ws);
  const range = rangeCreate(cells);
  const bounds = [range.top, range.left, range.bottom, range.right];
  const describe = (): string =>
    cells.map(cell => (typeof cell === "object" ? JSON.stringify(cell) : String(cell))).join(", ");
  if (
    !bounds.every(n => Number.isInteger(n) && n >= 1) ||
    range.bottom > MAX_ROW ||
    range.right > MAX_COL
  ) {
    throw new ExcelInvalidAddressError(
      describe(),
      "the range must name cells inside A1:XFD1048576",
      {
        field: "cells",
        where
      }
    );
  }
  if (
    range.sheetName !== undefined &&
    sheetNameKey(range.sheetName) !== sheetNameKey(getSheetName(ws))
  ) {
    throw new ExcelInvalidAddressError(
      describe(),
      `the range names sheet "${range.sheetName}", not this sheet ("${getSheetName(ws)}")`,
      { field: "cells", where }
    );
  }
  return range;
}

/**
 * Merge the cells of a range — an `A1:B2` string, its two corners, or four numbers — into one, keeping the top-left
 * cell's value and giving the others its style.
 *
 * A single cell is left as it is: there is nothing to merge, and Excel stores no such merge. It used to be recorded
 * and reported by `getMergedRegions`, then dropped when the workbook was written.
 *
 * A streaming worksheet writer (`Stream.WorksheetWriter`) is merged too, while the range's rows are still open.
 *
 * @throws {ExcelMergeConflictError} when the range overlaps a merge already on the sheet.
 * @throws {ExcelInvalidAddressError} when the range is outside the grid or names another sheet.
 * @throws {ExcelStreamStateError} when a streamed sheet has been committed.
 * @throws {ExcelRowOutOfBoundsError} when the range reaches a streamed row that has been committed.
 */
export function merge(ws: WorksheetData | StreamWorksheet, ...cells: RangeInput[]): void {
  const range = mergeRange("Worksheet.merge", ws, cells);
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._merge(range);
  } else if (range.top !== range.bottom || range.left !== range.right) {
    mergeCells(ws as WorksheetData, range.top, range.left, range.bottom, range.right);
  }
}

/** {@link merge}, leaving every cell's own style as it is. */
export function mergeWithoutStyle(
  ws: WorksheetData | StreamWorksheet,
  ...cells: RangeInput[]
): void {
  const range = mergeRange("Worksheet.mergeWithoutStyle", ws, cells);
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._merge(range, true);
  } else if (range.top !== range.bottom || range.left !== range.right) {
    mergeCellsWithoutStyle(ws as WorksheetData, range.top, range.left, range.bottom, range.right);
  }
}

/**
 * Unmerge every merged region that `cells` overlaps — an `A1:B2` range, its corners, or four numbers.
 *
 * @throws {ExcelUnsupportedFeatureError} for a streaming worksheet writer: a merge's cells may already be written.
 */
export function unmerge(ws: WorksheetData, ...cells: RangeInput[]): void {
  refuseStreamWorksheet(
    ws,
    "Worksheet.unmerge",
    "cannot unmerge: the merged cells may already have been written; merge only the ranges you mean to keep"
  );
  const range = mergeRange("Worksheet.unmerge", ws, cells);
  unMergeCells(ws, range.top, range.left, range.bottom, range.right);
}

/** Fill `range` with `formula`, written once on its top-left cell and shared to the rest (or as one array formula). */
export function fillFormula(
  ws: WorksheetData,
  range: string,
  formula: string,
  options?: FillFormulaOptions
): void;
/**
 * @deprecated Pass the results and the share type in the options object:
 * `fillFormula(ws, range, formula, { results, shareType })`. Removed in 1.0.
 */
export function fillFormula(
  ws: WorksheetData,
  range: string,
  formula: string,
  results?: FillFormulaOptions["results"],
  shareType?: string
): void;
export function fillFormula(
  ws: WorksheetData,
  range: string,
  formula: string,
  resultsOrOptions?: FillFormulaOptions | FillFormulaOptions["results"],
  shareType?: string
): void {
  checkWorksheet(ws);
  const check = argumentChecks("Worksheet.fillFormula");
  check.required("range", range);
  check.reference("range", range, { range: true });
  // A number is a formula too (`<f>5</f>`), and has always been accepted.
  if (typeof formula === "number") {
    check.number("formula", formula);
  } else {
    check.required("formula", formula);
    check.string("formula", formula, { nonEmpty: true });
  }
  if (typeof resultsOrOptions !== "function" && isRecordArgument(resultsOrOptions)) {
    checkSchema("Worksheet.fillFormula", "options", fillFormulaOptionsSchema, resultsOrOptions);
  }
  coreFillFormula(
    ws,
    normalizeReference(range),
    formula,
    resultsOrOptions as FillFormulaOptions["results"],
    shareType
  );
}

/**
 * Protect the sheet, optionally with a password: `protect(ws, { password: "pw", formatCells: true })`. A streaming
 * worksheet writer is protected too, until it is committed.
 */
export function protect(
  ws: WorksheetData | StreamWorksheet,
  options?: ProtectSheetOptions
): Promise<void>;
/**
 * @deprecated Pass the password in the options object: `protect(ws, { password, ...options })`.
 * Removed in 1.0.
 */
export function protect(
  ws: WorksheetData,
  password: string | undefined,
  options?: Partial<SheetProtection>
): Promise<void>;
export function protect(
  ws: WorksheetData | StreamWorksheet,
  passwordOrOptions?: string | ProtectSheetOptions,
  legacyOptions?: Partial<SheetProtection>
): Promise<void> {
  const writer = asStreamWorksheet(ws);
  try {
    checkWorksheet(ws);
    checkProtectCall("Worksheet.protect", passwordOrOptions, legacyOptions, "sheet");
    writer?._checkOpen("protect the worksheet");
  } catch (error) {
    return Promise.reject(error);
  }
  if (writer) {
    const [password, options] = splitPasswordOptions(passwordOrOptions, legacyOptions);
    return writer._protect(password, options);
  }
  return coreProtect(ws as WorksheetData, passwordOrOptions as string | undefined, legacyOptions);
}

/** Remove the sheet's protection, as {@link protect} set it — on a streaming worksheet writer too, until it is committed. */
export function unprotect(ws: WorksheetData | StreamWorksheet): void {
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._checkOpen("unprotect the worksheet");
    writer.sheetProtection = null;
    return;
  }
  coreUnprotect(ws as WorksheetData);
}

/** Where {@link addJson} and {@link addAoa} start, checked and resolved to a 1-based row and column. */
function resolveOrigin(
  where: string,
  ws: WorksheetData,
  origin: unknown
): { row: number; col: number } {
  const check = argumentChecks(where);
  if (origin === undefined) {
    return { row: 1, col: 1 };
  }
  if (typeof origin === "string") {
    check.reference("options.origin", origin);
    const cell = decodeCell(origin);
    return { row: cell.r + 1, col: cell.c + 1 };
  }
  if (typeof origin === "number") {
    if (origin === -1) {
      return { row: coreGetRowCount(ws) + 1, col: 1 };
    }
    check.number("options.origin", origin, { min: 0, max: MAX_ROW - 1, integer: true });
    return { row: origin + 1, col: 1 };
  }
  if (!check.record("options.origin", origin)) {
    return { row: 1, col: 1 };
  }
  check.number("options.origin.r", origin.r, { min: 0, max: MAX_ROW - 1, integer: true });
  check.number("options.origin.c", origin.c, { min: 0, max: MAX_COL - 1, integer: true });
  return { row: (origin.r as number) + 1, col: (origin.c as number) + 1 };
}

/** The block `rows` × `columns` from `start` must fit the grid, or part of it is written before the call fails. */
function checkExtent(
  where: string,
  start: { row: number; col: number },
  rows: number,
  columns: number
): void {
  const lastRow = start.row + rows - 1;
  const lastCol = start.col + columns - 1;
  if (rows > 0 && lastRow > MAX_ROW) {
    throw new ExcelRowOutOfBoundsError(lastRow, "the data reaches past the last row, 1048576", {
      field: "data",
      where
    });
  }
  if (columns > 0 && lastCol > MAX_COL) {
    throw new ExcelColumnOutOfBoundsError(
      lastCol,
      "the data reaches past the last column, XFD (16384)",
      {
        field: "data",
        where
      }
    );
  }
}

/**
 * Write a list of records as a table: a header row of their keys, then one row per record. Every record and value is
 * checked first, so a call that throws writes nothing.
 */
export function addJson(
  ws: WorksheetData,
  data: readonly Readonly<Record<string, CellValueInput>>[],
  options?: AddJsonOptions
): WorksheetData {
  const check = argumentChecks("Worksheet.addJson");
  check.required("data", data);
  check.array("data", data);
  const keys = new Set<string>();
  data.forEach((row, i) => {
    check.required(`data[${i}]`, row);
    if (check.record(`data[${i}]`, row)) {
      for (const [key, value] of Object.entries(row)) {
        keys.add(key);
        checkCellValue(value, { where: "Worksheet.addJson", field: `data[${i}].${key}` });
      }
    }
  });
  checkSchema("Worksheet.addJson", "options", addJsonOptionsSchema.optional(), options);
  options?.header?.forEach(key => keys.add(key));
  if (data.length > 0) {
    const start = resolveOrigin("Worksheet.addJson", ws, options?.origin);
    checkExtent("Worksheet.addJson", start, data.length + (options?.skipHeader ? 0 : 1), keys.size);
  }
  return addJSON(ws, data as readonly Record<string, CellValue>[], options);
}

/** Write rows of values from the origin (default `A1`). Every row and value is checked first, so a call that throws writes nothing. */
export function addAoa(
  ws: WorksheetData,
  data: readonly (readonly CellValueInput[])[],
  options?: AddAoaOptions
): WorksheetData {
  const check = argumentChecks("Worksheet.addAoa");
  check.required("data", data);
  check.array("data", data);
  let columns = 0;
  data.forEach((row, i) => {
    // A missing row is skipped, as it always was.
    if (row === undefined || row === null) {
      return;
    }
    check.array(`data[${i}]`, row);
    row.forEach((value, c) =>
      checkCellValue(value, { where: "Worksheet.addAoa", field: `data[${i}][${c}]` })
    );
    columns = Math.max(columns, row.length);
  });
  checkSchema("Worksheet.addAoa", "options", addAoaOptionsSchema.optional(), options);
  if (data.length > 0) {
    checkExtent(
      "Worksheet.addAoa",
      resolveOrigin("Worksheet.addAoa", ws, options?.origin),
      data.length,
      columns
    );
  }
  return addAOA(ws, data as readonly (readonly CellValue[])[], options);
}

/** A positional count as the deprecated forms coerce it (`Math.trunc`, negatives to 0): it must come out a number in the grid. */
function checkLegacyCount(check: ArgumentChecks, field: string, value: unknown, max: number): void {
  if (value === undefined || value === null) {
    return;
  }
  const count = Math.max(0, Math.trunc(value as number));
  if (!(count <= max)) {
    check.fail(
      field,
      `a count that is a number${max === Infinity ? "" : ` of at most ${max}`}`,
      value
    );
  }
}

/**
 * Freeze rows and columns so they stay in view while the rest scrolls: `freeze(ws, { columns: 1, rows: 1 })` pins
 * the first column and the header row. Zero (or omitting) either means that axis does not freeze, and
 * `freeze(ws, {})` is the same as {@link unfreeze}.
 *
 * @throws {ExcelArgumentError} when a count is not a whole number from 0 to the grid's last column or row — a
 * negative count was dropped and `NaN` written as a pane Excel repairs.
 */
export function freeze(ws: WorksheetData, options?: FreezePanesOptions): void;
/**
 * @deprecated Pass the counts in the options object: `freeze(ws, { columns, rows })`. Removed in 1.0.
 */
export function freeze(ws: WorksheetData, columns?: number, rows?: number): void;
export function freeze(
  ws: WorksheetData,
  columnsOrOptions?: number | FreezePanesOptions,
  legacyRows?: number
): void {
  const check = argumentChecks("Worksheet.freeze");
  if (isRecordArgument(columnsOrOptions)) {
    checkSchema("Worksheet.freeze", "options", freezeOptionsSchema, columnsOrOptions);
  } else {
    // The deprecated form keeps its coercion (`"2"` is 2, `true` is 1, a negative count is 0, `null` is 0) and
    // refuses only what it turned into a pane Excel repairs: no number at all, or one past the grid.
    checkLegacyCount(check, "columns", columnsOrOptions, MAX_COL - 1);
    checkLegacyCount(check, "rows", legacyRows, MAX_ROW - 1);
  }
  coreFreeze(ws, columnsOrOptions as number | undefined, legacyRows);
}

/**
 * Split the sheet into resizable panes at a point measured in **twentieths of a point**.
 *
 * @throws {ExcelArgumentError} when an offset is not a finite number of at least 0.
 */
export function split(ws: WorksheetData, options?: SplitPanesOptions): void;
/**
 * @deprecated Pass the offsets in the options object: `split(ws, { x, y })`. Removed in 1.0.
 */
export function split(ws: WorksheetData, x?: number, y?: number): void;
export function split(
  ws: WorksheetData,
  xOrOptions?: number | SplitPanesOptions,
  legacyY?: number
): void {
  const check = argumentChecks("Worksheet.split");
  if (isRecordArgument(xOrOptions)) {
    checkSchema("Worksheet.split", "options", splitOptionsSchema, xOrOptions);
  } else {
    // As `freeze`: the deprecated form keeps its coercion and refuses only an offset that is no number at all.
    checkLegacyCount(check, "x", xOrOptions, Infinity);
    checkLegacyCount(check, "y", legacyY, Infinity);
  }
  coreSplit(ws, xOrOptions as number | undefined, legacyY);
}

/**
 * Put an auto-filter over a range — an `A1:C4` string or a `{ from, to }` pair — or remove it with `null` or
 * `undefined` (what {@link getAutoFilter} answers for a sheet with none, so its result can be given back).
 *
 * @throws {ExcelInvalidAddressError} when the range is not a cell range of this sheet: a sheet-qualified or malformed
 * reference was written as given, and Excel repaired the file.
 */
export function setAutoFilter(ws: WorksheetData, range: AutoFilter | null | undefined): void {
  if (range === undefined) {
    coreSetAutoFilter(ws, null);
    return;
  }
  if (range !== null) {
    checkAutoFilter("Worksheet.setAutoFilter", "range", range);
  }
  // Stored canonical: `a1:c4` is `A1:C4`, which is what the writer puts in the file.
  coreSetAutoFilter(ws, normalizeAutoFilter(range));
}

/**
 * A copy of the sheet as its round-trip model — what the writers serialise and {@link setModel} applies. Editing it
 * changes nothing; apply a changed copy with {@link setModel}.
 *
 * It used to share the cells' value records, the views and the page setup with the sheet, so an edit to it changed
 * the sheet with none of the setters' checks.
 */
export function getModel(ws: WorksheetData): WorksheetModel {
  return copyPlain(getSheetModel(ws));
}

/**
 * Replace the sheet's contents with `value` — the inverse of {@link getModel}.
 *
 * The model is checked before anything changes: what the load walks must be there, every cell must be a value a
 * workbook can hold and sit in the row its address names, and a new name must be one the workbook can take. A model
 * that still fails part-way — a merge that conflicts with another, a table the workbook refuses — leaves the sheet
 * as it was. A new `name` renames the sheet as {@link setName} does, so formulas and defined names that refer to
 * it follow.
 *
 * @throws {ExcelArgumentError} when the model is malformed.
 * @throws {ExcelWorksheetNameError} when `name` is taken or not a name Excel accepts.
 */
export function setModel(ws: WorksheetData, value: WorksheetModel): void {
  checkWorksheetAttached(ws);
  checkWorksheetModel("Worksheet.setModel", "value", value);
  const from = getSheetName(ws);
  const to = value.name;
  if (to !== undefined && to !== from) {
    refuseLongSheetName(to);
    validateSheetName(getSheetWorkbook(ws), to, ws);
  }
  const before = getSheetModel(ws);
  try {
    setSheetModel(ws, { ...copyPlain(value), name: from });
  } catch (error) {
    setSheetModel(ws, before);
    throw error;
  }
  if (to !== undefined && to !== from) {
    setSheetName(ws, to);
  }
}

/**
 * @deprecated Without `base` the result is 1-based, while `Row.getValues` and `Column.getValues` default to 0 — so
 * state it: `{ base: 1 }` keeps this result, `{ base: 0 }` indexes as `Row` and `Column` do. Calling without
 * `options` is removed in 1.0.
 */
export function getValues(ws: WorksheetData): CellValue[][];
/**
 * The sheet's values as a sparse array of rows, indexed as `options.base` says: `{ base: 0 }` gives plain JavaScript
 * indexing — row 1 at `0`, column A at `0`, each row what `Row.getValues` returns — and `{ base: 1 }` the numbered
 * form, index `r` being row `r` and, within it, index `c` column `c`, index `0` unused. For a dense 0-based array use
 * {@link toAoa}. Creates nothing.
 */
export function getValues(ws: WorksheetData, options: ValuesOptions): CellValue[][];
export function getValues(ws: WorksheetData, options?: ValuesOptions): CellValue[][] {
  const check = argumentChecks("Worksheet.getValues");
  if (check.record("options", options) && options.base !== 0 && options.base !== 1) {
    check.fail("options.base", "0 or 1", options.base);
  }
  if (options?.base !== 0) {
    return getSheetValues(ws);
  }
  checkWorksheet(ws);
  const rows: CellValue[][] = [];
  for (const row of ws._rows) {
    if (row) {
      rows[row.number - 1] = rowGetValues(row);
    }
  }
  return rows;
}

export { coreGetMergedRegions as getMergedRegions };

/**
 * The used range — the smallest range covering every cell that holds a value — or `undefined` for a sheet with no
 * value, as `Row.getLast` and `Column.getLast` answer for one. A copy.
 *
 * (The deprecated `dimensions` answers an empty sheet with a range of zeros, which `Range.*` reads as `A1`.)
 */
export function getDimensions(ws: WorksheetData): RangeData | undefined {
  const dimensions = coreGetSheetDimensions(ws);
  return dimensions.top < 1 || dimensions.left < 1 ? undefined : dimensions;
}

/*
 * Panes, the auto-filter: `freeze` / `split` / `unfreeze` and `setAutoFilter` write them,
 * `getPanes` / `getAutoFilter` read them back.
 */

/**
 * The sheet's frozen or split panes, in the words {@link freeze} and {@link split} take — `{ state: "frozen", columns,
 * rows }` or `{ state: "split", x, y }` — or `undefined` when the sheet has none. Give it back to the matching setter
 * to set the same panes on another sheet:
 *
 * ```ts
 * const panes = Worksheet.getPanes(source);
 * if (panes?.state === "frozen") Worksheet.freeze(target, panes);
 * else if (panes?.state === "split") Worksheet.split(target, panes);
 * ```
 *
 * The rest of the view (zoom, grid lines, …) is in `ws.views`.
 */
export function getPanes(ws: WorksheetData): WorksheetPanes | undefined {
  // Read loosely: a file's view may say `frozenSplit`, which the published view union does not name.
  const view = corePanes(ws) as { state?: string; xSplit?: unknown; ySplit?: unknown } | undefined;
  if (view === undefined) {
    return undefined;
  }
  // `xSplit` / `ySplit` count columns and rows in a frozen view and twentieths of a point in a split one.
  const across = typeof view.xSplit === "number" ? view.xSplit : 0;
  const down = typeof view.ySplit === "number" ? view.ySplit : 0;
  if (view.state === "split") {
    return { state: "split", x: across, y: down };
  }
  return view.state === "frozen" || view.state === "frozenSplit"
    ? { state: "frozen", columns: across, rows: down }
    : undefined;
}

/**
 * The auto-filter range, or `undefined` when the sheet has none. It is a copy: change it with
 * {@link setAutoFilter}.
 */
export function getAutoFilter(ws: WorksheetData): AutoFilter | undefined {
  const filter = coreAutoFilter(ws);
  return filter === null || filter === undefined ? undefined : copyPlain(filter);
}

/**
 * A worksheet handle. `id`, `state`, `pageSetup`, `headerFooter`, `properties`, `views` and the other
 * settings in the published type are plain fields to read and assign, and `dataValidations` is the
 * registry `DataValidation.*` takes. What a namespace manages — tables, pivot tables, conditional
 * formats, form controls, the auto filter — is reached through it; those fields and the internal state are
 * marked `@deprecated` and leave the published type in 1.0.
 */
export type { WorksheetData as Handle } from "@excel/core/worksheet-core";

// --- deprecated names ---

/** @deprecated Use {@link getDimensions}. Removed in 1.0. */
export const dimensions = coreGetSheetDimensions;
/** @deprecated Use {@link getMergedRegions}. Removed in 1.0. */
export const mergedRegions = coreGetMergedRegions;
/** @deprecated Use {@link getPanes}. Removed in 1.0. */
export const panes = corePanes;
/** @deprecated Use {@link getAutoFilter}, which answers `undefined` rather than `null`. Removed in 1.0. */
export const autoFilter = coreAutoFilter;
/** @deprecated Use `Row.splice`. Removed in 1.0. */
export const spliceRows = coreSpliceRows;
/** @deprecated Use `Column.splice`. Removed in 1.0. */
export const spliceColumns = coreSpliceColumns;
/** @deprecated Use `Row.insert`. Removed in 1.0. */
export const insertRow = coreInsertRow;
/** @deprecated Use `Row.insertMany`. Removed in 1.0. */
export const insertRows = coreInsertRows;
/** @deprecated Use `Row.duplicate`. Removed in 1.0. */
export const duplicateRow = coreDuplicateRow;
/** @deprecated Use `Column.autoFit`. Removed in 1.0. */
export const autoFitColumn = coreAutoFitColumn;
/** @deprecated Use `Column.autoFit`. Removed in 1.0. */
export const autoFitColumns = coreAutoFitColumns;
/** @deprecated Use `Row.autoFit`. Removed in 1.0. */
export const autoFitRow = coreAutoFitRow;
/** @deprecated Use `Row.autoFit`. Removed in 1.0. */
export const autoFitRows = coreAutoFitRows;
/** @deprecated Use `ConditionalFormatting.add`. Removed in 1.0. */
export const addConditionalFormatting = addConditionalFormattingChecked;
/**
 * @deprecated Use `ConditionalFormatting.remove` for an entry, a position or a predicate, and
 * `ConditionalFormatting.clear` for every entry (what calling this without a filter does).
 * Removed in 1.0.
 */
export const removeConditionalFormatting = coreRemoveConditionalFormatting;
/** @deprecated Use `Column.getCount`. Removed in 1.0. */
export const columnCount = coreGetColumnCount;
/** @deprecated Use `Column.getActualCount`. Removed in 1.0. */
export const actualColumnCount = coreGetActualColumnCount;
/** @deprecated Use `Row.getCount`. Removed in 1.0. */
export const rowCount = coreGetRowCount;
/** @deprecated Use `Row.getActualCount`. Removed in 1.0. */
export const actualRowCount = coreGetActualRowCount;
/** @deprecated Use `Column.setDefinitions`. Removed in 1.0. */
export const setColumns = coreSetColumns;
/** @deprecated Use `Column.getDefinitions`. Removed in 1.0. */
export const columnDefinitions = coreGetColumnDefinitions;
/** @deprecated Use `Column.getDefinitions`. Removed in 1.0. */
export const columns = coreGetColumns;
/**
 * @deprecated Use `Column.getLast`, which returns the column **number** (or `undefined` for an
 * empty sheet) rather than a column view, and creates no column record. Removed in 1.0.
 */
export const lastColumn = coreGetLastColumn;
/**
 * @deprecated Use `Row.getLast`, which returns the row **number** rather than the row record.
 * Removed in 1.0.
 */
export const lastRow = coreGetLastRow;
/** @deprecated Use `Row.add`. Removed in 1.0. */
export const addRow = coreAddRow;
/** @deprecated Use `Row.addMany`. Removed in 1.0. */
export const addRows = coreAddRows;
/** @deprecated Use the `Row` and `Cell` namespaces. Removed in 1.0. */
export const getRow = coreGetRow;
/** @deprecated Use the `Row` and `Cell` namespaces. Removed in 1.0. */
export const getRows = coreGetRows;
/** @deprecated Use `Row.has`. Removed in 1.0. */
export const findRow = coreFindRow;
/** @deprecated Use `Row.has`. Removed in 1.0. */
export const findRows = coreFindRows;
/** @deprecated Use `Row.each`. Removed in 1.0. */
export const eachRow = coreEachRow;
/** @deprecated Use `Workbook.removeWorksheet(wb, ws)`, which this does. Removed in 1.0. */
export const destroy = coreDestroy;
