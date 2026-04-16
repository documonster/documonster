import { parse } from "@excel/calc/formula-parser";
import { tokenize } from "@excel/calc/formula-tokenizer";
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
// DefinedNames class
// ============================================================================

class DefinedNames {
  matrixMap: Record<string, CellMatrix>;
  /**
   * Formula-based defined names: name → formula expression string.
   * These are names that map to an expression (e.g. "LAMBDA(x,y,x+y)",
   * "{1,2;3,4}") rather than a cell/range reference.
   */
  formulaMap: Record<string, string>;
  /**
   * Opaque defined names: name → original text + optional localSheetId.
   * These are entries whose content we cannot parse (e.g. error values,
   * string literals, unknown constructs) but must preserve for round-trip
   * fidelity with the source XLSX file.
   */
  opaqueMap: Record<string, OpaqueEntry>;

  constructor() {
    this.matrixMap = {};
    this.formulaMap = {};
    this.opaqueMap = {};
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
    delete this.formulaMap[name];
    delete this.opaqueMap[name];
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
    delete this.matrixMap[name];
    delete this.opaqueMap[name];
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
    Object.entries(this.matrixMap).forEach(([name, matrix]) => {
      matrix.forEach((cell: DefinedNameCell) => {
        callback(name, cell);
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
      .map(([name, matrix]) => matrix.findCellEx(address, false) && name)
      .filter((name): name is string => Boolean(name));
  }

  _explore(matrix: CellMatrix, cell: DefinedNameCell): Range {
    cell.mark = false;
    const sheetName = cell.sheetName!; // Always set for cells in defined names

    const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
    let x: number;
    let y: number;

    // Helper to get cell with proper type
    const getCell = (row: number, col: number): DefinedNameCell | undefined => {
      return matrix.findCellAt(sheetName, row, col) as DefinedNameCell | undefined;
    };

    // grow vertical - only one col to worry about
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

    // grow horizontal - ensure all rows can grow
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

  getRanges(name: string, matrix?: CellMatrix): DefinedNameModel {
    // Formula-based name takes precedence if no cell matrix exists
    const formula = this.formulaMap[name];
    matrix = matrix || this.matrixMap[name];

    if (!matrix) {
      if (formula) {
        return { name, ranges: [formula], formulaExpression: formula };
      }
      return { name, ranges: [] };
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
      name,
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
      .map(([name, matrix]) => this.getRanges(name, matrix))
      .filter((definedName: DefinedNameModel) => definedName.ranges.length);

    // Formula-based names from formulaMap (only include names not already in matrixMap)
    const formulaNames = Object.entries(this.formulaMap)
      .filter(([name]) => !this.matrixMap[name])
      .map(([name, expression]) => ({
        name,
        ranges: [expression],
        formulaExpression: expression
      }));

    // Opaque names — rawText preserved for round-trip
    const opaqueNames: DefinedNameModel[] = Object.entries(this.opaqueMap).map(([name, entry]) => ({
      name,
      ranges: [],
      rawText: entry.rawText,
      localSheetId: entry.localSheetId,
      kind: "opaque" as const
    }));

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
    const opaqueMap = (this.opaqueMap = {} as Record<string, OpaqueEntry>);

    for (const definedName of value) {
      // Programmatic API path: formulaExpression already set and no rawText
      if (definedName.formulaExpression && definedName.rawText === undefined) {
        formulaMap[definedName.name] = definedName.formulaExpression;
        continue;
      }

      // XLSX parse path (rawText present) or programmatic path with ranges only
      const classified = classifyDefinedName(definedName.rawText, definedName.ranges);

      switch (classified.kind) {
        case "reference": {
          const matrix = (matrixMap[definedName.name] = new CellMatrix());
          for (const rangeStr of classified.ranges) {
            if (rangeRegexp.test(rangeStr.split("!").pop() ?? "")) {
              matrix.addCell(rangeStr);
            }
          }
          break;
        }
        case "formula":
          formulaMap[definedName.name] = classified.formulaExpression!;
          break;
        case "opaque":
          if (definedName.rawText) {
            opaqueMap[definedName.name] = {
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
