import { Range } from "@excel/range";
import type { Address } from "@excel/types";
import { CellMatrix } from "@excel/utils/cell-matrix";
import { colCache, type DecodedRange } from "@excel/utils/col-cache";
import { getDefaultSyntaxProbe, type SyntaxProbe } from "@formula/default-syntax-probe";

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
 *
 * **Probe semantics:** `probe` is the formula tokenizer+parser oracle. It
 * is the *only* authority for deciding whether a non-range, non-wrapper
 * string is a formula. When `probe` is `null` (no formula engine
 * installed and no probe injected), any such string is classified as
 * **opaque** — we have no evidence it is a formula, and leaving it
 * opaque preserves round-trip bytes via `rawText`.
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
  const extracted = extractRanges(rawText);
  if (extracted.length > 0) {
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
   * Tracks the `hidden` flag for each storage key. Used by chartEx
   * `_xlchart.vN.M` infrastructure names — Excel emits them with
   * `hidden="1"` to suppress the Name Manager UI. A missing key
   * means "not hidden" (the default).
   */
  hiddenMap: Record<string, boolean>;
  /**
   * Opaque defined names: storageKey → original text + optional localSheetId.
   */
  opaqueMap: Record<string, OpaqueEntry>;
  /**
   * Reverse mapping: storageKey → original name (bare, without scope suffix).
   * Needed because storageKey encodes the scope, but consumers need the bare name.
   */
  nameForKey: Record<string, string>;

  /**
   * Optional explicit formula-syntax probe. When set, this is used to
   * classify non-range, non-wrapper defined-name text during `set model`.
   * When unset, the classifier falls back to the process-wide default
   * probe (set by `installFormulaEngine()`). When neither is available,
   * classification is conservative — non-range text becomes opaque.
   */
  private readonly _explicitProbe: SyntaxProbe | null;

  /**
   * @param probe Optional formula-syntax probe used when classifying
   *   defined-name text. Injecting a probe here makes classification
   *   deterministic for this instance regardless of process-global
   *   `installFormulaEngine()` state. When omitted, the instance defers
   *   to the default probe at classification time (see `set model`).
   */
  constructor(probe?: SyntaxProbe) {
    this.matrixMap = {};
    this.formulaMap = {};
    this.localSheetIdMap = {};
    this.hiddenMap = {};
    this.opaqueMap = {};
    this.nameForKey = {};
    this._explicitProbe = probe ?? null;
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
  addHidden(locStr: string, name: string): void {
    this.add(locStr, name);
    this.hiddenMap[name] = true;
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
        const hidden = this.hiddenMap[sKey];
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
          result.localSheetId = localSheetId;
        }
        if (this.hiddenMap[sKey]) {
          result.hidden = true;
        }
        return result;
      });

    // Opaque names — rawText preserved for round-trip
    const opaqueNames: DefinedNameModel[] = Object.entries(this.opaqueMap).map(([sKey, entry]) => {
      const bareName = this.nameForKey[sKey] ?? sKey;
      const out: DefinedNameModel = {
        name: bareName,
        ranges: [],
        rawText: entry.rawText,
        localSheetId: entry.localSheetId,
        kind: "opaque" as const
      };
      if (this.hiddenMap[sKey]) {
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
  set model(value: DefinedNameModel[]) {
    const matrixMap = (this.matrixMap = {} as Record<string, CellMatrix>);
    const formulaMap = (this.formulaMap = {} as Record<string, string>);
    const localSheetIdMap = (this.localSheetIdMap = {} as Record<string, number>);
    const hiddenMap = (this.hiddenMap = {} as Record<string, boolean>);
    const opaqueMap = (this.opaqueMap = {} as Record<string, OpaqueEntry>);
    const nameForKeyMap = (this.nameForKey = {} as Record<string, string>);

    // Resolve probe lazily: a caller may have constructed the Workbook
    // before `installFormulaEngine()` but load XLSX data afterwards. We
    // want whichever probe is registered at *load* time, not construct
    // time. An explicit per-instance probe always wins when provided.
    const probe = this._explicitProbe ?? getDefaultSyntaxProbe();

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
