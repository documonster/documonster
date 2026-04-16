import { parse } from "@excel/formula/syntax/parser";
import { tokenize } from "@excel/formula/syntax/tokenizer";
import { Range } from "@excel/range";
import type { Address } from "@excel/types";
import { CellMatrix } from "@excel/utils/cell-matrix";
import { colCache, type DecodedRange } from "@excel/utils/col-cache";

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

// Cell type for defined names - extends Address with mark for exploration algorithm
interface DefinedNameCell {
  sheetName?: string;
  address: string;
  row: number;
  col: number;
  mark?: boolean;
}

// Location can be a single cell address or a range
type CellLocation = Address | DecodedRange;

interface DefinedNameModel {
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

function isValidRange(range: string): boolean {
  // Skip array constants wrapped in {} - these are not valid cell ranges
  if (range.startsWith("{") || range.endsWith("}")) {
    return false;
  }

  // Extract the cell reference part (after the sheet name if present)
  const cellRef = range.split("!").pop() ?? "";

  // Must match one of the valid patterns
  if (
    !cellRangeRegexp.test(cellRef) &&
    !rowRangeRegexp.test(cellRef) &&
    !colRangeRegexp.test(cellRef)
  ) {
    return false;
  }

  try {
    const decoded = colCache.decodeEx(range);
    if (
      ("row" in decoded && typeof decoded.row === "number") ||
      ("top" in decoded && typeof decoded.top === "number") ||
      ("left" in decoded && typeof decoded.left === "number")
    ) {
      return true;
    }
    return false;
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
 */
function classifyDefinedName(
  rawText: string | undefined,
  ranges: string[]
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

  // If the text contains an unquoted parenthesis (e.g. OFFSET(...), LAMBDA(...)),
  // it is likely a formula — attempt formula parsing first so that extractRanges
  // does not mis-split the function arguments on commas.
  if (hasUnquotedParen(trimmed)) {
    // Skip array constants ({…}), string literals ("…"), and error values (#…)
    // which may contain parentheses but are not formulas.
    const isArrayConst = trimmed.startsWith("{") && trimmed.endsWith("}");
    const isStringLit = trimmed.startsWith('"') && trimmed.endsWith('"');
    const isErrorVal = trimmed.startsWith("#");

    if (!isArrayConst && !isStringLit && !isErrorVal) {
      try {
        const tokens = tokenize(trimmed);
        if (tokens.length > 0) {
          parse(tokens);
          return { kind: "formula", ranges: [trimmed], formulaExpression: trimmed };
        }
      } catch {
        // Parse failed — fall through to extractRanges / opaque
      }
    }
  }

  // Try to extract cell/range references (handles comma-separated multi-area names).
  const extracted = extractRanges(rawText);
  if (extracted.length > 0) {
    return { kind: "reference", ranges: extracted };
  }

  // For content that is clearly not a useful formula (array constants, error
  // values, string literals), classify as opaque even if the parser could
  // technically handle them. These should be preserved verbatim for round-trip
  // but not participate in calculation.
  const isArrayConst = trimmed.startsWith("{") && trimmed.endsWith("}");
  const isStringLit = trimmed.startsWith('"') && trimmed.endsWith('"');
  const isErrorVal = trimmed.startsWith("#");
  if (isArrayConst || isStringLit || isErrorVal) {
    return { kind: "opaque", ranges: [] };
  }

  // No unquoted parens — still try formula parser for simple constant expressions.
  if (!hasUnquotedParen(trimmed)) {
    try {
      const tokens = tokenize(trimmed);
      if (tokens.length > 0) {
        parse(tokens);
        return { kind: "formula", ranges: [trimmed], formulaExpression: trimmed };
      }
    } catch {
      // Parse failed
    }
  }

  // Nothing worked — opaque
  return { kind: "opaque", ranges: [] };
}

// ============================================================================
// Internal Key — disambiguates same-name entries with different scopes
// ============================================================================

/**
 * Build the internal storage key for a defined name entry.
 * Workbook-scoped: just the bare name.
 * Sheet-scoped: `"name\0sheetId"` (null char separator avoids collisions).
 */
function storageKey(name: string, localSheetId?: number): string {
  return localSheetId !== undefined ? `${name}\0${localSheetId}` : name;
}

// ============================================================================
// DefinedNames class
// ============================================================================

class DefinedNames {
  matrixMap: Record<string, CellMatrix>;
  /**
   * Formula-based defined names: storageKey → formula expression string.
   */
  formulaMap: Record<string, string>;
  /**
   * Tracks the localSheetId for each storage key.
   * If a key is not in this map, the name is workbook-scoped (global).
   */
  localSheetIdMap: Record<string, number>;
  /**
   * Opaque defined names: storageKey → original text + optional localSheetId.
   */
  opaqueMap: Record<string, OpaqueEntry>;
  /**
   * Reverse mapping: storageKey → original name (bare, without scope suffix).
   * Needed because storageKey encodes the scope, but consumers need the bare name.
   */
  nameForKey: Record<string, string>;

  constructor() {
    this.matrixMap = {};
    this.formulaMap = {};
    this.localSheetIdMap = {};
    this.opaqueMap = {};
    this.nameForKey = {};
  }

  getMatrix(name: string): CellMatrix {
    const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
    return matrix;
  }

  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add(locStr: string, name: string): void {
    const location = colCache.decodeEx(locStr);
    if ("error" in location) {
      return; // Invalid reference, skip
    }
    // A name is either cell-reference or formula — clear any formula/opaque binding
    // (programmatic API always operates on workbook-scoped names)
    delete this.formulaMap[name];
    delete this.opaqueMap[name];
    this.nameForKey[name] = name;
    this.addEx(location, name);
  }

  addEx(location: CellLocation, name: string): void {
    const matrix = this.getMatrix(name);
    if ("top" in location) {
      // It's a range (DecodedRange has top/left/bottom/right from Location)
      for (let col = location.left; col <= location.right; col++) {
        for (let row = location.top; row <= location.bottom; row++) {
          const address = {
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row,
            col
          };

          matrix.addCellEx(address);
        }
      }
    } else {
      // It's a single cell address
      matrix.addCellEx(location);
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
  addFormula(name: string, expression: string): void {
    // A name is either formula or cell-reference — clear any cell-reference/opaque binding
    // (programmatic API always operates on workbook-scoped names)
    delete this.matrixMap[name];
    delete this.opaqueMap[name];
    this.nameForKey[name] = name;
    this.formulaMap[name] = expression;
  }

  remove(locStr: string, name: string): void {
    const location = colCache.decodeEx(locStr);
    if ("error" in location) {
      return; // Invalid reference, skip
    }
    this.removeEx(location, name);
  }

  removeEx(location: CellLocation, name: string): void {
    const matrix = this.getMatrix(name);
    if ("top" in location) {
      // Range - remove each cell
      for (let col = location.left; col <= location.right; col++) {
        for (let row = location.top; row <= location.bottom; row++) {
          matrix.removeCellEx({
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row,
            col
          });
        }
      }
    } else {
      matrix.removeCellEx(location);
    }
  }

  removeAllNames(location: CellLocation): void {
    Object.values(this.matrixMap).forEach((matrix: CellMatrix) => {
      if ("top" in location) {
        // Range - remove each cell
        for (let col = location.left; col <= location.right; col++) {
          for (let row = location.top; row <= location.bottom; row++) {
            matrix.removeCellEx({
              sheetName: location.sheetName,
              address: colCache.n2l(col) + row,
              row,
              col
            });
          }
        }
      } else {
        matrix.removeCellEx(location);
      }
    });
  }

  forEach(callback: (name: string, cell: DefinedNameCell) => void): void {
    Object.entries(this.matrixMap).forEach(([sKey, matrix]) => {
      const bareName = this.nameForKey[sKey] ?? sKey;
      matrix.forEach((cell: DefinedNameCell) => {
        callback(bareName, cell);
      });
    });
  }

  // get all the names of a cell
  getNames(addressStr: string): string[] {
    const location = colCache.decodeEx(addressStr);
    if ("error" in location || "top" in location) {
      return []; // Invalid reference or range not supported
    }
    return this.getNamesEx(location);
  }

  getNamesEx(address: Address): string[] {
    return Object.entries(this.matrixMap)
      .map(([sKey, matrix]) => matrix.findCellEx(address, false) && (this.nameForKey[sKey] ?? sKey))
      .filter((name): name is string => Boolean(name));
  }

  /**
   * Return all defined name entries in this collection, including scope info.
   * Each entry has the bare name and optional localSheetId.
   * Same bare name may appear multiple times with different scopes.
   */
  getAllNames(): { name: string; localSheetId?: number }[] {
    return this.getAllEntries().map(e =>
      e.localSheetId !== undefined
        ? { name: e.name, localSheetId: e.localSheetId }
        : { name: e.name }
    );
  }

  /**
   * Return all defined name entries with full details (name, ranges, scope).
   *
   * This is the primary enumeration API. Each entry is self-contained —
   * no second lookup is needed. Same bare name may appear multiple times
   * with different scopes.
   */
  getAllEntries(): DefinedNameModel[] {
    const result: DefinedNameModel[] = [];
    const seen = new Set<string>();
    for (const sKey of Object.keys(this.matrixMap)) {
      if (seen.has(sKey)) {
        continue;
      }
      seen.add(sKey);
      const model = this.getRanges(sKey);
      const localSheetId = this.localSheetIdMap[sKey];
      result.push(localSheetId !== undefined ? { ...model, localSheetId } : model);
    }
    for (const sKey of Object.keys(this.formulaMap)) {
      if (seen.has(sKey)) {
        continue;
      }
      seen.add(sKey);
      const model = this.getRanges(sKey);
      const localSheetId = this.localSheetIdMap[sKey];
      result.push(localSheetId !== undefined ? { ...model, localSheetId } : model);
    }
    for (const sKey of Object.keys(this.opaqueMap)) {
      if (seen.has(sKey)) {
        continue;
      }
      seen.add(sKey);
      const bareName = this.nameForKey[sKey] ?? sKey;
      const entry = this.opaqueMap[sKey];
      result.push({
        name: bareName,
        ranges: [],
        rawText: entry.rawText,
        localSheetId: entry.localSheetId,
        kind: "opaque" as const
      });
    }
    return result;
  }

  _explore(matrix: CellMatrix, cell: DefinedNameCell): Range {
    cell.mark = false;
    const sheetName = cell.sheetName!;

    const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
    let x: number;
    let y: number;

    const getCell = (row: number, col: number): DefinedNameCell | undefined => {
      return matrix.findCellAt(sheetName, row, col) as DefinedNameCell | undefined;
    };

    function vGrow(yy: number, edge: "top" | "bottom"): boolean {
      const c = getCell(yy, cell.col);
      if (!c || !c.mark) {
        return false;
      }
      range[edge] = yy;
      c.mark = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, "top"); y--) {
      /* advance */
    }
    for (y = cell.row + 1; vGrow(y, "bottom"); y++) {
      /* advance */
    }

    function hGrow(xx: number, edge: "left" | "right"): boolean {
      const cells: DefinedNameCell[] = [];
      for (y = range.top; y <= range.bottom; y++) {
        const c = getCell(y, xx);
        if (c && c.mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      range[edge] = xx;
      for (let i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, "left"); x--) {
      /* advance */
    }
    for (x = cell.col + 1; hGrow(x, "right"); x++) {
      /* advance */
    }

    return range;
  }

  /**
   * Get ranges for a specific scoped entry.
   *
   * Unlike `getRanges(name)` which uses the bare name (and may hit the
   * wrong scope when the same name exists both globally and locally),
   * this method uses the internal `storageKey` to look up the exact entry.
   */
  getRangesScoped(name: string, localSheetId?: number): DefinedNameModel {
    const sKey = storageKey(name, localSheetId);
    return this.getRanges(sKey);
  }

  getRanges(name: string, matrix?: CellMatrix): DefinedNameModel {
    // `name` can be a bare name (backward compat) or a storageKey.
    // Try storageKey first, then bare name (workbook-scoped).
    const formula = this.formulaMap[name];
    matrix = matrix || this.matrixMap[name];
    const bareName = this.nameForKey[name] ?? name;

    if (!matrix) {
      if (formula) {
        return { name: bareName, ranges: [formula], formulaExpression: formula };
      }
      return { name: bareName, ranges: [] };
    }

    // mark and sweep!
    matrix.forEach((cell: DefinedNameCell) => {
      cell.mark = true;
    });
    const ranges = matrix
      .map((cell: DefinedNameCell) => cell.mark && this._explore(matrix!, cell))
      .filter((range): range is Range => Boolean(range))
      .map((range: Range) => range.$shortRange);

    return {
      name: bareName,
      ranges
    };
  }

  normaliseMatrix(matrix: CellMatrix, sheetName: string): void {
    // some of the cells might have shifted on specified sheet
    // need to reassign rows, cols
    matrix.forEachInSheet(
      sheetName,
      (cell: DefinedNameCell | undefined, row: number, col: number) => {
        if (cell) {
          if (cell.row !== row || cell.col !== col) {
            cell.row = row;
            cell.col = col;
            cell.address = colCache.n2l(col) + row;
          }
        }
      }
    );
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    Object.values(this.matrixMap).forEach((matrix: CellMatrix) => {
      matrix.spliceRows(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    });
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    Object.values(this.matrixMap).forEach((matrix: CellMatrix) => {
      matrix.spliceColumns(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    });
  }

  get model(): DefinedNameModel[] {
    // Cell-reference based names from matrixMap
    const cellNames = Object.entries(this.matrixMap)
      .map(([sKey, matrix]) => {
        const result = this.getRanges(sKey, matrix);
        const localSheetId = this.localSheetIdMap[sKey];
        if (localSheetId !== undefined) {
          return { ...result, localSheetId };
        }
        return result;
      })
      .filter((definedName: DefinedNameModel) => definedName.ranges.length);

    // Formula-based names from formulaMap (only include names not already in matrixMap)
    const formulaNames = Object.entries(this.formulaMap)
      .filter(([sKey]) => !this.matrixMap[sKey])
      .map(([sKey, expression]) => {
        const bareName = this.nameForKey[sKey] ?? sKey;
        const result: DefinedNameModel = {
          name: bareName,
          ranges: [expression],
          formulaExpression: expression
        };
        const localSheetId = this.localSheetIdMap[sKey];
        if (localSheetId !== undefined) {
          return { ...result, localSheetId };
        }
        return result;
      });

    // Opaque names — rawText preserved for round-trip
    const opaqueNames: DefinedNameModel[] = Object.entries(this.opaqueMap).map(([sKey, entry]) => {
      const bareName = this.nameForKey[sKey] ?? sKey;
      return {
        name: bareName,
        ranges: [],
        rawText: entry.rawText,
        localSheetId: entry.localSheetId,
        kind: "opaque" as const
      };
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
  set model(value: DefinedNameModel[]) {
    const matrixMap = (this.matrixMap = {} as Record<string, CellMatrix>);
    const formulaMap = (this.formulaMap = {} as Record<string, string>);
    const localSheetIdMap = (this.localSheetIdMap = {} as Record<string, number>);
    const opaqueMap = (this.opaqueMap = {} as Record<string, OpaqueEntry>);
    const nameForKeyMap = (this.nameForKey = {} as Record<string, string>);

    for (const definedName of value) {
      const sKey = storageKey(definedName.name, definedName.localSheetId);
      nameForKeyMap[sKey] = definedName.name;

      // Track localSheetId for all name kinds
      if (definedName.localSheetId !== undefined) {
        localSheetIdMap[sKey] = definedName.localSheetId;
      }

      // Programmatic API path: formulaExpression already set and no rawText
      if (definedName.formulaExpression && definedName.rawText === undefined) {
        formulaMap[sKey] = definedName.formulaExpression;
        continue;
      }

      // XLSX parse path (rawText present) or programmatic path with ranges only
      const classified = classifyDefinedName(definedName.rawText, definedName.ranges);

      switch (classified.kind) {
        case "reference": {
          const matrix = (matrixMap[sKey] = new CellMatrix());
          for (const rangeStr of classified.ranges) {
            if (rangeRegexp.test(rangeStr.split("!").pop() ?? "")) {
              matrix.addCell(rangeStr);
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
}

export { DefinedNames, type DefinedNameModel };
