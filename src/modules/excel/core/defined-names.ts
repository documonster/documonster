import {
  cellRegionAdd,
  cellRegionContains,
  cellRegionRects,
  cellRegionEdit,
  cellRegionIsEmpty,
  cellRegionRemove,
  cellRegionRenameSheet,
  createCellRegion,
  type CellRect,
  type CellRegion
} from "@excel/core/cell-region";
import { rangeAbsoluteShort, rangeCreate, serialisedSheetName } from "@excel/core/range";
import { definedNameProblem } from "@excel/core/table-name";
import { DefinedNameError, ExcelArgumentError } from "@excel/errors";
import type { DecodedAddress } from "@excel/types";
import type { DecodedRange } from "@excel/utils/col-cache";
import { colCache } from "@excel/utils/col-cache";
import type { ReferenceEdit } from "@excel/utils/reference-rewrite";
import { REF_ERROR, rewriteFormula } from "@excel/utils/reference-rewrite";
import { excelNameKey, sheetNameKey } from "@utils/name-key";

/**
 * A formula-syntax oracle: returns `true` iff `text` parses as a formula
 * expression. Used to classify defined-name text (formula vs. opaque) when
 * loading XLSX. Consumers may inject a custom probe via
 * `createWorkbook({ formulaSyntaxProbe })` / `createDefinedNames(probe)`;
 * otherwise the caller's fallback is used — the built-in tokenizer + parser
 * probe in `core/model-load.ts`, bound by every public entry and by the
 * readers, so no install / registration step is required.
 */
export type SyntaxProbe = (text: string) => boolean;

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;
const wholeAxisRegexp = /^(?:[$][A-Za-z]{1,3}:[$][A-Za-z]{1,3}|[$]\d+:[$]\d+)$/;
const absoluteAreaRegexp = /^[$][A-Za-z]{1,3}[$]\d+(?::[$][A-Za-z]{1,3}[$]\d+)?$/;

/**
 * True when `range` names a fixed set of cells: every coordinate absolute (`Sheet!$A$1:$B$2`, `$A:$A`, `$1:$1`), and
 * a sheet named whenever a `!` is present. `!$A$1` is Excel's form for "the sheet the name is used on" — it has no
 * sheet, and written back as `$A$1` it is a different definition.
 */
function isRegionReference(range: string): boolean {
  const bang = range.lastIndexOf("!");
  if (bang === 0) {
    return false;
  }
  const reference = range.slice(bang + 1);
  return absoluteAreaRegexp.test(reference) || wholeAxisRegexp.test(reference);
}

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

// Location can be a single cell address or a range
type CellLocation = DecodedAddress | DecodedRange;

/** A defined name as persisted: its key plus the ranges it covers. */
export interface DefinedNameModel {
  name: string;
  ranges: string[];
  localSheetId?: number;
  /** Formula expression for formula-based defined names (e.g. "LAMBDA(x,y,x+y)") */
  formulaExpression?: string;
  /** Original XML text — preserved for opaque names round-trip */
  rawText?: string;
  /**
   * Classification determined by the semantic layer:
   * - "reference": pure cell/range union — stored in matrixMap
   * - "formula": parseable expression — stored in formulaMap
   * - "opaque": unrecognised content preserved for round-trip — stored in opaqueMap
   */
  kind?: "reference" | "formula" | "opaque";
  /**
   * `definedName/@hidden` — when `true`, emits `hidden="1"` on the
   * XML element. Used by chartEx `_xlchart.vN.M` defined names that
   * Excel creates automatically for every chartEx data reference;
   * these names are infrastructure, never intended for users, so
   * Excel hides them from the Name Manager UI.
   */
  hidden?: boolean;
}

/** Stored entry for an opaque (unrecognised) defined name. */
interface OpaqueEntry {
  rawText: string;
  localSheetId?: number;
}

// ============================================================================
// Range validation helpers (moved from defined-name-xform.ts)
// ============================================================================

// Regex to validate cell range format:
// - Cell: $A$1 or A1
// - Range: $A$1:$B$10 or A1:B10
// - Row range: $1:$2 (for print titles)
// - Column range: $A:$B (for print titles)
const cellRangeRegexp = /^[$]?[A-Za-z]{1,3}[$]?\d+(:[$]?[A-Za-z]{1,3}[$]?\d+)?$/;
const rowRangeRegexp = /^[$]?\d+:[$]?\d+$/;
const colRangeRegexp = /^[$]?[A-Za-z]{1,3}:[$]?[A-Za-z]{1,3}$/;

/**
 * An unquoted sheet name as a formula writes one — an identifier, as the formula tokenizer reads a word — or a 3-D
 * pair, optionally after an external workbook's `[1]`.
 */
const UNQUOTED_SHEET =
  /^(?:\[[^\]]*\])?[\p{L}_\\][\p{L}\p{M}\p{N}_.]*(?::[\p{L}_\\][\p{L}\p{M}\p{N}_.]*)?$/u;
/** A quoted one, `''` standing for a quote. */
const QUOTED_SHEET = /^'(?:[^']|'')+'$/;

/**
 * Whether `prefix` (what comes before a reference's `!`) is a sheet name a formula can write. `@S!$A$1:$A$3` and
 * `-S!$A$1` are formulas whose operand is on `S`, not regions on sheets named `@S` and `-S`.
 */
function isSheetPrefix(prefix: string): boolean {
  return QUOTED_SHEET.test(prefix) || UNQUOTED_SHEET.test(prefix);
}

function isValidRange(range: string): boolean {
  // Skip array constants wrapped in {} - these are not valid cell ranges
  if (range.startsWith("{") || range.endsWith("}")) {
    return false;
  }

  // Extract the cell reference part (after the sheet name if present)
  const cellRef = range.split("!").pop() ?? "";
  const bang = range.lastIndexOf("!");
  if (bang > 0 && !isSheetPrefix(range.slice(0, bang))) {
    return false;
  }

  // Must match one of the valid patterns
  if (
    !cellRangeRegexp.test(cellRef) &&
    !rowRangeRegexp.test(cellRef) &&
    !colRangeRegexp.test(cellRef)
  ) {
    return false;
  }

  try {
    return locationRect(colCache.decodeEx(range)) !== undefined;
  } catch {
    return false;
  }
}

/**
 * Extract valid cell/range references from a raw defined name text string.
 * Handles comma-separated ranges and quoted sheet names that may contain commas.
 */
function extractRanges(parsedText: string): string[] {
  const trimmed = parsedText.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return [];
  }

  const ranges: string[] = [];
  let quotesOpened = false;
  let last = "";
  parsedText.split(",").forEach(item => {
    if (!item) {
      return;
    }
    const quotes = (item.match(/'/g) ?? []).length;

    if (!quotes) {
      if (quotesOpened) {
        last += `${item},`;
      } else if (isValidRange(item)) {
        ranges.push(item);
      }
      return;
    }
    const quotesEven = quotes % 2 === 0;

    if (!quotesOpened && quotesEven && isValidRange(item)) {
      ranges.push(item);
    } else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      if (isValidRange(last + item)) {
        ranges.push(last + item);
      }
      last = "";
    } else {
      quotesOpened = true;
      last += `${item},`;
    }
  });
  return ranges;
}

// ============================================================================
// Classifier — Stage 2 of the two-phase defined name design
// ============================================================================

/**
 * Check whether a string contains a '(' character outside of single-quoted
 * sheet name segments. Sheet names in cell references use single quotes:
 * e.g. `'Sheet (1)'!$A$1` — the '(' is inside quotes and does NOT indicate
 * a function call. A genuine formula like `OFFSET(Sheet1!$A$1,0,0,3,1)`
 * has '(' outside of any quotes.
 */
function hasUnquotedParen(s: string): boolean {
  let inQuote = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'") {
      inQuote = !inQuote;
    } else if (ch === "(" && !inQuote) {
      return true;
    }
  }
  return false;
}

/**
 * Classify a raw defined name text into one of three categories:
 * 1. **reference** — pure cell/range union (goes into matrixMap)
 * 2. **formula** — parseable expression (goes into formulaMap)
 * 3. **opaque** — unrecognised content preserved for round-trip
 *
 * Classification order matters:
 * - If the text contains an unquoted `(`, it is likely a formula — try the
 *   formula parser first (this prevents `extractRanges` from misinterpreting
 *   function arguments as partial range references).
 * - Otherwise try to extract ranges (pure cell references).
 * - If neither works, fall back to opaque.
 *
 * **Probe semantics:** `probe` is the formula tokenizer+parser oracle. It
 * is the *only* authority for deciding whether a non-range, non-wrapper
 * string is a formula. When `probe` is `null`, any such string is classified
 * as **opaque** — we have no evidence it is a formula, and leaving it opaque
 * preserves round-trip bytes via `rawText`. (In practice `definedNamesSetModel`
 * always supplies a probe: its caller's fallback is required.)
 *
 * This function is pure: the classification of a given input depends
 * entirely on (rawText, ranges, probe) and no global state. Two calls
 * with the same arguments always produce the same result.
 */
function classifyDefinedName(
  rawText: string | undefined,
  ranges: string[],
  probe: SyntaxProbe | null
): { kind: "reference" | "formula" | "opaque"; ranges: string[]; formulaExpression?: string } {
  // If rawText is missing, fall back to existing ranges (programmatic API path)
  if (rawText === undefined) {
    if (ranges.length > 0) {
      return { kind: "reference", ranges };
    }
    return { kind: "opaque", ranges: [] };
  }

  const trimmed = rawText.trim();
  if (trimmed.length === 0) {
    return { kind: "opaque", ranges: [] };
  }

  // Opaque-looking wrappers — array constants, string literals, error
  // values — are preserved verbatim and never routed through formula
  // parsing. Detect them once, up front.
  const isArrayConst = trimmed.startsWith("{") && trimmed.endsWith("}");
  const isStringLit = trimmed.startsWith('"') && trimmed.endsWith('"');
  const isErrorVal = trimmed.startsWith("#");
  const isOpaqueWrapper = isArrayConst || isStringLit || isErrorVal;

  const hasParen = hasUnquotedParen(trimmed);

  // If the text contains an unquoted parenthesis (e.g. OFFSET(...),
  // LAMBDA(...)), it is either a formula or a malformed expression —
  // never a cell-range union. `extractRanges` is not safe to call on
  // such text because it splits on commas and can mis-identify
  // `OFFSET(Sheet1` as a partial range reference.
  //
  // With a probe: confirm the expression parses before classifying as
  // formula; if the probe rejects it, fall through to opaque.
  // Without a probe: we cannot confirm it parses, so preserve the text
  // verbatim as opaque. This is deliberately strict — the alternative
  // (silently promoting unverified text to `formula`) produced
  // classification results that depended on global install state.
  if (hasParen && !isOpaqueWrapper) {
    if (probe && probe(trimmed)) {
      return { kind: "formula", ranges: [trimmed], formulaExpression: trimmed };
    }
    return { kind: "opaque", ranges: [] };
  }

  // Try to extract cell/range references (handles comma-separated
  // multi-area names). Safe to call only after the paren check above,
  // because `extractRanges` splits on commas.
  //
  // **Only a fully absolute union is a region.** A relative or mixed reference in a name (`S!A1`, `S!$A1:$A3`)
  // resolves against the cell that uses it, so it names no fixed cells: modelled as a region it came out empty and
  // the name was dropped on write, and read as `$A$1` it would change what every formula using it computes. Nor may
  // an item be skipped — the union has to account for the whole text, or the rest is lost. Anything else stays text.
  const extracted = extractRanges(rawText);
  if (
    extracted.length > 0 &&
    extracted.join(",") === trimmed &&
    extracted.every(isRegionReference)
  ) {
    return { kind: "reference", ranges: extracted };
  }

  // Opaque wrappers that didn't pass as a formula above classify
  // straight to opaque — we preserve them verbatim without calculating.
  if (isOpaqueWrapper) {
    return { kind: "opaque", ranges: [] };
  }

  // No parens and not a range — could still be a parseable constant
  // expression or a reference to another defined name. Only classify as
  // formula if a probe can confirm it parses; otherwise stay opaque so
  // round-trip text is preserved without silently promoting unparseable
  // content to the formula bucket.
  if (probe && probe(trimmed)) {
    return { kind: "formula", ranges: [trimmed], formulaExpression: trimmed };
  }

  // Nothing worked — opaque
  return { kind: "opaque", ranges: [] };
}

// ============================================================================
// Internal Key — disambiguates same-name entries with different scopes
// ============================================================================

/**
 * Build the internal storage key for a defined name entry: the name as Excel compares it (`excelNameKey` — `Rate` and
 * `RATE`, or `ß` and `SS`, are one name; `ı` and `I` are two), and for a sheet-scoped name `"key\0sheetId"` (null
 * char separator avoids collisions). The name as written is kept in `nameForKey`.
 */
function storageKey(name: string, localSheetId?: number): string {
  const key = excelNameKey(name);
  return localSheetId !== undefined ? `${key}\0${localSheetId}` : key;
}

/**
 * Refuse a name the workbook already uses for a table. Excel's UI refuses it, and a file holding both is ambiguous:
 * a formula's `Name` could mean either. Only the programmatic entry points check; a model load keeps what a file holds.
 */
function refuseTableName(dn: DefinedNamesData, name: string): void {
  if (dn._tableNames?.has(excelNameKey(name)) === true) {
    throw new DefinedNameError(
      `Defined name "${name}" is the name of a table in this workbook. ` +
        `Tables and defined names share one namespace, compared as Excel compares them (case-insensitive).`
    );
  }
}

/** Whether `dn` defines `name`, in any scope, compared as Excel compares names. */
export function definedNamesHasName(dn: DefinedNamesData, name: string): boolean {
  const key = excelNameKey(name);
  return definedNamesModel(dn).some(entry => excelNameKey(entry.name) === key);
}

/**
 * Refuse a name a caller chose that Excel cannot hold — one that reads as a cell reference (`A1`), starts with a
 * digit, holds a space — or that a table of the workbook already uses. The programmatic entry points check before
 * they change anything; a model load keeps what a file holds.
 */
export function definedNamesCheckName(dn: DefinedNamesData, where: string, name: unknown): void {
  const problem = definedNameProblem(name);
  if (problem !== undefined) {
    throw new DefinedNameError(
      `${where}: ${JSON.stringify(name)} is not a valid name — ${problem}.`
    );
  }
  refuseTableName(dn, name as string);
}

/**
 * Give the cell at `location` the workbook-scoped name `name`, as {@link definedNamesAdd} does for a reference string:
 * a name is either a reference or a formula, so a formula or opaque binding of the same name is replaced. Leaving it
 * in place shadowed it behind the reference — and removing the cell again then dropped the name from the file.
 */
export function definedNamesAddCell(
  dn: DefinedNamesData,
  location: CellLocation,
  name: string
): void {
  refuseTableName(dn, name);
  const key = storageKey(name);
  delete dn.formulaMap[key];
  delete dn.opaqueMap[key];
  definedNamesAddEx(dn, location, name);
}

/** A storage key as given, or the workbook-scoped key of a bare name. */
function keyOf(nameOrKey: string): string {
  return nameOrKey.includes("\0") ? nameOrKey : storageKey(nameOrKey);
}

// ============================================================================
// DefinedNames class
// ============================================================================

// ============================================================================
// DefinedNames — de-classed domain model (data record + flat helpers)
// ============================================================================

/**
 * Plain-data defined-names registry (de-classed domain model). All former
 * methods are flat `definedNames*` helpers taking the record as first arg.
 */
export interface DefinedNamesData {
  /** The cells each reference-defined name covers, keyed by storage key. */
  matrixMap: Record<string, CellRegion>;
  formulaMap: Record<string, string>;
  localSheetIdMap: Record<string, number>;
  hiddenMap: Record<string, boolean>;
  opaqueMap: Record<string, OpaqueEntry>;
  nameForKey: Record<string, string>;
  /** Optional per-instance formula-syntax probe (see createDefinedNames). */
  _explicitProbe: SyntaxProbe | null;
  /**
   * The `excelNameKey`s of the owning workbook's tables, when the registry belongs to one. Tables and defined names
   * share one namespace, so `add` / `addFormula` refuse a name in it.
   */
  _tableNames?: ReadonlySet<string>;
}

/**
 * Refuse something that is not a defined-names registry (`Workbook.getDefinedNames(wb)`) where one is required, with
 * an `ExcelArgumentError` naming `field`.
 */
export function checkDefinedNames(dn: unknown, field = "dn"): asserts dn is DefinedNamesData {
  if ((dn as { matrixMap?: unknown } | null | undefined)?.matrixMap === undefined) {
    throw new ExcelArgumentError(
      field,
      `${field} must be a defined-names registry (Workbook.getDefinedNames(wb)), got ${
        dn === null || typeof dn !== "object" ? String(dn) : "an object that is not one"
      }.`
    );
  }
}

/** Create a defined-names registry, optionally with an explicit syntax probe. */
export function createDefinedNames(probe?: SyntaxProbe): DefinedNamesData {
  return {
    matrixMap: {},
    formulaMap: {},
    localSheetIdMap: {},
    hiddenMap: {},
    opaqueMap: {},
    nameForKey: {},
    _explicitProbe: probe ?? null
  };
}

/**
 * The rectangle a decoded location covers, or `undefined` when it is not a location on the sheet.
 *
 * A whole-column reference (`$A:$B`) spans every row and a whole-row one (`$1:$3`) every column. A bound past the
 * sheet's 1,048,576 rows or 16,384 columns makes the location invalid rather than large: it names nothing, and
 * accepting it is how a name came to cover more cells than any sheet has.
 */
function locationRect(location: CellLocation | { error: string }): CellRect | undefined {
  if ("error" in location) {
    return undefined;
  }
  let rect: CellRect;
  if ("top" in location) {
    // `colCache` leaves the missing axis of `$A:$B` or `$1:$3` unset (`NaN` or `null`).
    const wholeColumns = !Number.isFinite(location.top) && !Number.isFinite(location.bottom);
    const wholeRows = !Number.isFinite(location.left) && !Number.isFinite(location.right);
    rect = {
      top: wholeColumns ? 1 : location.top,
      bottom: wholeColumns ? MAX_ROWS : location.bottom,
      left: wholeRows ? 1 : location.left,
      right: wholeRows ? MAX_COLUMNS : location.right
    };
  } else {
    rect = { top: location.row, bottom: location.row, left: location.col, right: location.col };
  }
  const inGrid = (value: number, max: number): boolean =>
    Number.isInteger(value) && value >= 1 && value <= max;
  if (
    !inGrid(rect.top, MAX_ROWS) ||
    !inGrid(rect.bottom, MAX_ROWS) ||
    !inGrid(rect.left, MAX_COLUMNS) ||
    !inGrid(rect.right, MAX_COLUMNS)
  ) {
    return undefined;
  }
  return {
    top: Math.min(rect.top, rect.bottom),
    bottom: Math.max(rect.top, rect.bottom),
    left: Math.min(rect.left, rect.right),
    right: Math.max(rect.left, rect.right)
  };
}

/**
 * Whether {@link definedNamesAdd} can place a name at `locStr` — a cell or a rectangle of the grid, sheet-qualified.
 * The adder itself skips one it cannot, which is what a reader of a file wants; a caller handing one over gets this
 * answer as an error instead.
 */
export function definedNamesIsPlaceable(locStr: string): boolean {
  if (typeof locStr !== "string") {
    return false;
  }
  try {
    return locationRect(colCache.decodeEx(locStr)) !== undefined;
  } catch {
    return false;
  }
}

function regionFor(dn: DefinedNamesData, name: string): CellRegion {
  return Object.hasOwn(dn.matrixMap, name)
    ? dn.matrixMap[name]
    : (dn.matrixMap[name] = createCellRegion());
}

// add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
/** `localSheetId` scopes the name to the sheet at that 0-based tab position; omitted, the name is the workbook's. */
export function definedNamesAdd(
  dn: DefinedNamesData,
  locStr: string,
  name: string,
  localSheetId?: number
): void {
  const location = colCache.decodeEx(locStr);
  if (locationRect(location) === undefined || "error" in location) {
    return; // Invalid reference, skip
  }
  refuseTableName(dn, name);
  // A name is either cell-reference or formula — clear any formula/opaque binding in the same scope
  const key = storageKey(name, localSheetId);
  delete dn.formulaMap[key];
  delete dn.opaqueMap[key];
  dn.nameForKey[key] = name;
  definedNamesAddEx(dn, location, name, localSheetId);
}

/**
 * Register a hidden defined name that maps a chartEx `_xlchart.vN.M`
 * pointer to a worksheet range. Excel emits one of these for every
 * data reference in a chartEx chart:
 *
 *   <definedName name="_xlchart.v1.0" hidden="1">Sheet1!$A$1:$A$3</definedName>
 *
 * The chartEx then references `_xlchart.v1.0` from its `<cx:f>`
 * element instead of the worksheet range directly. Direct sheet
 * references in `<cx:f>` are rejected by Excel 2016+ with
 * "Removed Part: /xl/drawings/drawingN.xml (Drawing shape)" on
 * load, so chartEx data MUST go through this indirection.
 *
 * Same semantics as `add()` but also marks the name `hidden` so
 * it does not show up in Excel's Name Manager UI.
 *
 * @param locStr - Worksheet reference (e.g. `"Sheet1!$A$1:$A$3"`)
 * @param name   - Defined name to register (e.g. `"_xlchart.v1.0"`)
 */
export function definedNamesAddHidden(dn: DefinedNamesData, locStr: string, name: string): void {
  definedNamesAdd(dn, locStr, name);
  dn.hiddenMap[storageKey(name)] = true;
}

export function definedNamesAddEx(
  dn: DefinedNamesData,
  location: CellLocation,
  name: string,
  localSheetId?: number
): void {
  const key = storageKey(name, localSheetId);
  dn.nameForKey[key] ??= name;
  if (localSheetId !== undefined) {
    dn.localSheetIdMap[key] = localSheetId;
  }
  const region = regionFor(dn, key);
  const rect = locationRect(location);
  if (rect) {
    cellRegionAdd(region, location.sheetName, rect);
  }
}

/**
 * Register a formula-based defined name.
 *
 * Unlike `add()` which binds a name to a cell/range reference, this binds
 * a name to an arbitrary formula expression that will be evaluated at
 * calculation time.
 *
 * @param name - The defined name (e.g. "MyArray")
 * @param expression - The formula expression (e.g. "{1,2;3,4}", "LAMBDA(x,y,x+y)")
 */
export function definedNamesAddFormula(
  dn: DefinedNamesData,
  name: string,
  expression: string,
  localSheetId?: number
): void {
  refuseTableName(dn, name);
  // A name is either formula or cell-reference — clear any cell-reference/opaque binding in the same scope
  const key = storageKey(name, localSheetId);
  delete dn.matrixMap[key];
  delete dn.opaqueMap[key];
  dn.nameForKey[key] = name;
  dn.formulaMap[key] = expression;
  if (localSheetId !== undefined) {
    dn.localSheetIdMap[key] = localSheetId;
  }
}

/**
 * Delete the name `name` in the scope `localSheetId` (the workbook when omitted) — reference, formula or opaque, all of
 * it. Returns `true` if the registry had it. Formulas that use the name are left as they are and read `#NAME?`, as they
 * do when Excel's Name Manager deletes one.
 */
export function definedNamesDelete(
  dn: DefinedNamesData,
  name: string,
  localSheetId?: number
): boolean {
  const key = storageKey(name, localSheetId);
  const had =
    Object.hasOwn(dn.matrixMap, key) ||
    Object.hasOwn(dn.formulaMap, key) ||
    Object.hasOwn(dn.opaqueMap, key);
  delete dn.matrixMap[key];
  delete dn.formulaMap[key];
  delete dn.opaqueMap[key];
  delete dn.hiddenMap[key];
  delete dn.localSheetIdMap[key];
  delete dn.nameForKey[key];
  return had;
}

/** Remove `name` from the cells at `locStr`. Returns `true` if any cell carried it. */
export function definedNamesRemove(dn: DefinedNamesData, locStr: string, name: string): boolean {
  const location = colCache.decodeEx(locStr);
  if ("error" in location) {
    return false; // Invalid reference, skip
  }
  return definedNamesRemoveEx(dn, location, name);
}

export function definedNamesRemoveEx(
  dn: DefinedNamesData,
  location: CellLocation,
  name: string
): boolean {
  // Looked up, not created: an empty cell-reference entry would shadow a formula-defined name of the same name,
  // which `definedNamesModel` then drops — removing a name nothing carried deleted the formula name from the file.
  const key = storageKey(name);
  const region = Object.hasOwn(dn.matrixMap, key) ? dn.matrixMap[key] : undefined;
  const rect = locationRect(location);
  if (!region || !rect) {
    return false;
  }
  return cellRegionRemove(region, location.sheetName, rect);
}

export function definedNamesRemoveAllNames(dn: DefinedNamesData, location: CellLocation): void {
  const rect = locationRect(location);
  if (!rect) {
    return;
  }
  for (const region of Object.values(dn.matrixMap)) {
    cellRegionRemove(region, location.sheetName, rect);
  }
}

// get all the names of a cell
/** The names that cover the cell at `addressStr` (`"Sheet1!A1"`); `[]` for an invalid reference or a range. */
export function definedNamesGetNames(dn: DefinedNamesData, addressStr: string): string[] {
  const location = colCache.decodeEx(addressStr);
  if ("error" in location || "top" in location) {
    return []; // Invalid reference or range not supported
  }
  return definedNamesGetNamesEx(dn, location);
}

export function definedNamesGetNamesEx(dn: DefinedNamesData, address: DecodedAddress): string[] {
  return Object.entries(dn.matrixMap)
    .filter(([, region]) => cellRegionContains(region, address.sheetName, address.row, address.col))
    .map(([sKey]) => dn.nameForKey[sKey] ?? sKey);
}

/**
 * Return all defined name entries in this collection, including scope info.
 * Each entry has the bare name and optional localSheetId.
 * Same bare name may appear multiple times with different scopes.
 */
export function definedNamesGetAllNames(
  dn: DefinedNamesData
): { name: string; localSheetId?: number }[] {
  return definedNamesGetAllEntries(dn).map(e =>
    e.localSheetId !== undefined ? { name: e.name, localSheetId: e.localSheetId } : { name: e.name }
  );
}

/**
 * Return all defined name entries with full details (name, ranges, scope).
 *
 * This is the primary enumeration API. Each entry is self-contained —
 * no second lookup is needed. Same bare name may appear multiple times
 * with different scopes.
 */
export function definedNamesGetAllEntries(dn: DefinedNamesData): DefinedNameModel[] {
  const result: DefinedNameModel[] = [];
  const seen = new Set<string>();
  // `hidden` is part of what an entry is: leaving it off made a hidden name read as a visible one.
  const withHidden = (sKey: string, entry: DefinedNameModel): DefinedNameModel =>
    dn.hiddenMap[sKey] ? { ...entry, hidden: true } : entry;
  for (const sKey of Object.keys(dn.matrixMap)) {
    if (seen.has(sKey)) {
      continue;
    }
    seen.add(sKey);
    const model = definedNamesGetRanges(dn, sKey);
    const localSheetId = dn.localSheetIdMap[sKey];
    result.push(withHidden(sKey, localSheetId !== undefined ? { ...model, localSheetId } : model));
  }
  for (const sKey of Object.keys(dn.formulaMap)) {
    if (seen.has(sKey)) {
      continue;
    }
    seen.add(sKey);
    const model = definedNamesGetRanges(dn, sKey);
    const localSheetId = dn.localSheetIdMap[sKey];
    result.push(withHidden(sKey, localSheetId !== undefined ? { ...model, localSheetId } : model));
  }
  for (const sKey of Object.keys(dn.opaqueMap)) {
    if (seen.has(sKey)) {
      continue;
    }
    seen.add(sKey);
    const bareName = dn.nameForKey[sKey] ?? sKey;
    const entry = dn.opaqueMap[sKey];
    result.push(
      withHidden(sKey, {
        name: bareName,
        ranges: [],
        rawText: entry.rawText,
        localSheetId: entry.localSheetId,
        kind: "opaque" as const
      })
    );
  }
  return result;
}

/**
 * A rectangle as a name's range text. One spanning every row is a whole-column reference and one spanning every
 * column a whole-row reference, as Excel shows and stores them — `$A:$A`, not `$A$1:$A$1048576`. The whole sheet is
 * written as rows, as Excel writes it.
 */
function rectText(sheetName: string | undefined, rect: CellRect): string {
  const range =
    sheetName === undefined
      ? rangeCreate(rect.top, rect.left, rect.bottom, rect.right)
      : rangeCreate(rect.top, rect.left, rect.bottom, rect.right, sheetName);
  if (rect.left === 1 && rect.right === MAX_COLUMNS) {
    return `${serialisedSheetName(range)}$${rect.top}:$${rect.bottom}`;
  }
  if (rect.top === 1 && rect.bottom === MAX_ROWS) {
    return `${serialisedSheetName(range)}$${colCache.n2l(rect.left)}:$${colCache.n2l(rect.right)}`;
  }
  return rangeAbsoluteShort(range);
}

/**
 * Get ranges for a specific scoped entry.
 *
 * Unlike `getRanges(name)` which uses the bare name (and may hit the
 * wrong scope when the same name exists both globally and locally),
 * this method uses the internal `storageKey` to look up the exact entry.
 */
export function definedNamesGetRangesScoped(
  dn: DefinedNamesData,
  name: string,
  localSheetId?: number
): DefinedNameModel {
  const sKey = storageKey(name, localSheetId);
  return definedNamesGetRanges(dn, sKey);
}

export function definedNamesGetRanges(dn: DefinedNamesData, name: string): DefinedNameModel {
  // `name` can be a bare name (backward compat) or a storageKey.
  // Try storageKey first, then bare name (workbook-scoped).
  const key = keyOf(name);
  const formula = dn.formulaMap[key];
  const region = Object.hasOwn(dn.matrixMap, key) ? dn.matrixMap[key] : undefined;
  const bareName = dn.nameForKey[key] ?? name;

  if (!region) {
    if (formula) {
      return { name: bareName, ranges: [formula], formulaExpression: formula };
    }
    return { name: bareName, ranges: [] };
  }

  return {
    name: bareName,
    ranges: cellRegionRects(region).map(({ sheet, rect }) => rectText(sheet, rect))
  };
}

/**
 * Carry every name through a structural edit — rows or columns inserted or deleted, a sheet renamed or removed — by
 * the rules in `reference-rewrite`.
 *
 * A reference-defined name's areas move, grow and shrink with the cells; a formula-defined name's expression is
 * rewritten like any formula. A reference name whose every area the edit deletes keeps its name and refers to
 * `#REF!`, as Excel's does: dropping it would turn every formula that uses it from `#REF!` into `#NAME?`.
 * Opaque names are text this library could not read, and are left as they are.
 */
export function definedNamesApplyEdit(dn: DefinedNamesData, edit: ReferenceEdit): void {
  for (const key of Object.keys(dn.matrixMap)) {
    const region = dn.matrixMap[key];
    const hadCells = !cellRegionIsEmpty(region);
    if (edit.kind === "rename") {
      dn.matrixMap[key] = cellRegionRenameSheet(region, edit.from, edit.to);
    } else if (edit.kind === "removeSheet") {
      dn.matrixMap[key] = cellRegionRenameSheet(region, edit.sheet, undefined);
    } else {
      cellRegionEdit(region, edit);
    }
    if (hadCells && cellRegionIsEmpty(dn.matrixMap[key])) {
      delete dn.matrixMap[key];
      dn.formulaMap[key] = REF_ERROR;
    }
  }
  for (const [key, expression] of Object.entries(dn.formulaMap)) {
    dn.formulaMap[key] = rewriteFormula(expression, edit);
  }
  if (edit.kind === "removeSheet") {
    const position = edit.sheets.findIndex(name => sheetNameKey(name) === sheetNameKey(edit.sheet));
    if (position >= 0) {
      removeSheetScope(dn, position);
    }
  }
}

/**
 * Drop the names scoped to the sheet at tab position `position` and renumber the scopes after it.
 *
 * A sheet-scoped name's `localSheetId` is the sheet's position on the tab bar, chartsheets included, so removing a
 * sheet both deletes the names that belonged to it — as Excel does — and moves every later scope down by one. Left
 * alone, each of those names would have been read back as belonging to the next sheet along.
 */
function removeSheetScope(dn: DefinedNamesData, position: number): void {
  const localSheetIds = dn.localSheetIdMap;
  const opaque = dn.opaqueMap;
  const names = dn.nameForKey;
  const scopeOf = (key: string): number | undefined =>
    localSheetIds[key] ?? opaque[key]?.localSheetId;
  const rekey = (key: string): string | undefined => {
    const scope = scopeOf(key);
    if (scope === undefined || scope < position) {
      return key;
    }
    if (scope === position) {
      return undefined;
    }
    return storageKey(names[key] ?? key.slice(0, key.indexOf("\0")), scope - 1);
  };
  const remap = <T>(map: Record<string, T>, change?: (value: T, key: string) => T) => {
    const next: Record<string, T> = {};
    for (const [key, value] of Object.entries(map)) {
      const newKey = rekey(key);
      if (newKey !== undefined) {
        next[newKey] = change ? change(value, key) : value;
      }
    }
    return next;
  };
  const shifted = (key: string): number | undefined => {
    const scope = scopeOf(key);
    return scope === undefined || scope < position ? scope : scope - 1;
  };
  dn.matrixMap = remap(dn.matrixMap);
  dn.formulaMap = remap(dn.formulaMap);
  dn.hiddenMap = remap(dn.hiddenMap);
  dn.opaqueMap = remap(dn.opaqueMap, (entry, key) => ({ ...entry, localSheetId: shifted(key) }));
  dn.localSheetIdMap = remap(dn.localSheetIdMap, (_, key) => shifted(key)!);
  dn.nameForKey = remap(dn.nameForKey);
}

export function definedNamesModel(dn: DefinedNamesData): DefinedNameModel[] {
  // Cell-reference based names from matrixMap
  const cellNames = Object.entries(dn.matrixMap)
    .map(([sKey]) => {
      const result = definedNamesGetRanges(dn, sKey);
      const localSheetId = dn.localSheetIdMap[sKey];
      const hidden = dn.hiddenMap[sKey];
      const out: DefinedNameModel = { ...result };
      if (localSheetId !== undefined) {
        out.localSheetId = localSheetId;
      }
      if (hidden) {
        out.hidden = true;
      }
      return out;
    })
    .filter((definedName: DefinedNameModel) => definedName.ranges.length);

  // Formula-based names from formulaMap (only include names not already in matrixMap)
  const formulaNames = Object.entries(dn.formulaMap)
    .filter(([sKey]) => !dn.matrixMap[sKey])
    .map(([sKey, expression]) => {
      const bareName = dn.nameForKey[sKey] ?? sKey;
      const result: DefinedNameModel = {
        name: bareName,
        ranges: [expression],
        formulaExpression: expression
      };
      const localSheetId = dn.localSheetIdMap[sKey];
      if (localSheetId !== undefined) {
        result.localSheetId = localSheetId;
      }
      if (dn.hiddenMap[sKey]) {
        result.hidden = true;
      }
      return result;
    });

  // Opaque names — rawText preserved for round-trip
  const opaqueNames: DefinedNameModel[] = Object.entries(dn.opaqueMap).map(([sKey, entry]) => {
    const bareName = dn.nameForKey[sKey] ?? sKey;
    const out: DefinedNameModel = {
      name: bareName,
      ranges: [],
      rawText: entry.rawText,
      localSheetId: entry.localSheetId,
      kind: "opaque" as const
    };
    if (dn.hiddenMap[sKey]) {
      out.hidden = true;
    }
    return out;
  });

  return [...cellNames, ...formulaNames, ...opaqueNames];
}

/**
 * Deserialise an array of `DefinedNameModel` entries (typically from XLSX parsing).
 *
 * Stage 2 of the two-phase design: each entry's `rawText` is classified
 * into reference / formula / opaque by `classifyDefinedName()`.  Entries
 * that arrive without `rawText` (programmatic API) fall back to inspecting
 * the existing `ranges` and `formulaExpression` fields for compatibility.
 */
export function definedNamesSetModel(
  dn: DefinedNamesData,
  value: readonly DefinedNameModel[],
  fallbackProbe: SyntaxProbe
): void {
  const matrixMap = (dn.matrixMap = {} as Record<string, CellRegion>);
  const formulaMap = (dn.formulaMap = {} as Record<string, string>);
  const localSheetIdMap = (dn.localSheetIdMap = {} as Record<string, number>);
  const hiddenMap = (dn.hiddenMap = {} as Record<string, boolean>);
  const opaqueMap = (dn.opaqueMap = {} as Record<string, OpaqueEntry>);
  const nameForKeyMap = (dn.nameForKey = {} as Record<string, string>);

  // Resolve the probe: an explicit per-instance probe always wins; otherwise
  // the caller's fallback. It is a parameter rather than an import so that the
  // tokenizer and parser belong to whoever loads a model — see `core/model-load.ts`.
  const probe = dn._explicitProbe ?? fallbackProbe;

  for (const definedName of value) {
    const sKey = storageKey(definedName.name, definedName.localSheetId);
    nameForKeyMap[sKey] = definedName.name;

    // Track localSheetId for all name kinds
    if (definedName.localSheetId !== undefined) {
      localSheetIdMap[sKey] = definedName.localSheetId;
    }
    if (definedName.hidden) {
      hiddenMap[sKey] = true;
    }

    // Programmatic API path: formulaExpression already set and no rawText
    if (definedName.formulaExpression && definedName.rawText === undefined) {
      formulaMap[sKey] = definedName.formulaExpression;
      continue;
    }

    // XLSX parse path (rawText present) or programmatic path with ranges only
    const classified = classifyDefinedName(definedName.rawText, definedName.ranges, probe);

    switch (classified.kind) {
      case "reference": {
        const region = (matrixMap[sKey] = createCellRegion());
        for (const rangeStr of classified.ranges) {
          const reference = rangeStr.split("!").pop() ?? "";
          if (!rangeRegexp.test(reference) && !wholeAxisRegexp.test(reference)) {
            continue;
          }
          const location = colCache.decodeEx(rangeStr);
          const rect = locationRect(location);
          if (rect) {
            cellRegionAdd(region, location.sheetName, rect);
          }
        }
        break;
      }
      case "formula":
        formulaMap[sKey] = classified.formulaExpression!;
        break;
      case "opaque":
        if (definedName.rawText) {
          opaqueMap[sKey] = {
            rawText: definedName.rawText,
            localSheetId: definedName.localSheetId
          };
        }
        break;
    }
  }
}
