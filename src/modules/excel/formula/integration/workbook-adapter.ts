/**
 * Workbook Adapter — Bridge between live workbook objects and the snapshot.
 *
 * This is the **only** file in the new engine pipeline that imports live
 * workbook/worksheet/cell types. All other engine code depends exclusively
 * on the snapshot types.
 *
 * ## Responsibilities
 *
 * 1. `buildWorkbookSnapshot()` — walk the live workbook and produce an
 *    immutable `WorkbookSnapshot`.
 * 2. Cell value conversion — Date → serial number, rich text → string,
 *    shared formula translation, etc.
 */

import type { Cell } from "@excel/cell";
import { Enums } from "@excel/enums";
import type { Worksheet } from "@excel/worksheet";
import { dateToExcel } from "@utils/utils.base";

import type { WorkbookLike } from "../materialize/types";
import type {
  CalcPropertiesSnapshot,
  CellSnapshot,
  DefinedNameSnapshot,
  FormulaCellKind,
  ResolvedTable,
  SnapshotCellValue,
  SnapshotErrorValue,
  TableColumnSnapshot,
  TableSnapshot,
  WorkbookPropertiesSnapshot,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "./workbook-snapshot";
import { snapshotCellKey, scopedNameKey } from "./workbook-snapshot";

// ============================================================================
// Build Workbook Snapshot
// ============================================================================

/**
 * Build a complete `WorkbookSnapshot` from a live workbook.
 *
 * This traverses all worksheets and cells once, converting everything to
 * engine-internal snapshot types. The result is a fully self-contained,
 * read-only data structure.
 */
export function buildWorkbookSnapshot(workbook: WorkbookLike): WorkbookSnapshot {
  const worksheets: WorksheetSnapshot[] = [];
  const worksheetsByName = new Map<string, WorksheetSnapshot>();
  const worksheetsById = new Map<number, WorksheetSnapshot>();

  const date1904 = workbook.properties?.date1904 ?? false;

  for (const ws of workbook.worksheets) {
    const wsSnapshot = buildWorksheetSnapshot(ws, date1904);
    worksheets.push(wsSnapshot);
    worksheetsByName.set(ws.name.toLowerCase(), wsSnapshot);
    worksheetsById.set(ws.id, wsSnapshot);
  }

  const definedNames = buildDefinedNames(workbook);

  // Build table-by-name index for O(1) lookup
  const tablesByName = new Map<string, ResolvedTable>();
  for (const wsSnapshot of worksheets) {
    for (const table of wsSnapshot.tables) {
      if (table.name) {
        tablesByName.set(table.name.toLowerCase(), {
          table,
          sheetName: wsSnapshot.name
        });
      }
    }
  }

  const calcProperties: CalcPropertiesSnapshot = {
    fullCalcOnLoad: workbook.calcProperties?.fullCalcOnLoad,
    iterate: workbook.calcProperties?.iterate,
    iterateCount: workbook.calcProperties?.iterateCount,
    iterateDelta: workbook.calcProperties?.iterateDelta
  };

  const properties: WorkbookPropertiesSnapshot = {
    date1904
  };

  return {
    worksheets,
    worksheetsByName,
    worksheetsById,
    definedNames,
    tablesByName,
    calcProperties,
    properties
  };
}

// ============================================================================
// Build Worksheet Snapshot
// ============================================================================

function buildWorksheetSnapshot(ws: Worksheet, date1904: boolean): WorksheetSnapshot {
  const cells = new Map<string, CellSnapshot>();
  const hiddenRows = new Set<number>();

  // Use includeEmpty so we observe the `hidden` flag on rows that have
  // no populated cells — a user may hide an empty row (e.g. filter UI)
  // and SUBTOTAL(1xx,…) still needs to treat that row as hidden.
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (row.hidden) {
      hiddenRows.add(rowNumber);
    }
    row.eachCell((cell, colNumber) => {
      const cellSnapshot = buildCellSnapshot(cell, rowNumber, colNumber, date1904);
      if (cellSnapshot) {
        cells.set(snapshotCellKey(rowNumber, colNumber), cellSnapshot);
      }
    });
  });

  const dims = ws.dimensions;
  const dimensions = dims
    ? { top: dims.top, left: dims.left, bottom: dims.bottom, right: dims.right }
    : null;

  const tables = buildTables(ws);

  return {
    id: ws.id,
    name: ws.name,
    dimensions,
    cells,
    hiddenRows,
    tables
  };
}

// ============================================================================
// Build Cell Snapshot
// ============================================================================

function buildCellSnapshot(
  cell: Cell,
  row: number,
  col: number,
  date1904: boolean
): CellSnapshot | null {
  const cellType = cell.type;

  // Skip truly empty cells
  if (cellType === Enums.ValueType.Null) {
    return null;
  }

  // ── Formula cells ──
  if (cellType === Enums.ValueType.Formula) {
    return buildFormulaCellSnapshot(cell, row, col, date1904);
  }

  // ── Non-formula cells ──
  const value = convertCellValue(cell.value, date1904);

  return {
    row,
    col,
    value,
    formulaKind: "none"
  };
}

function buildFormulaCellSnapshot(
  cell: Cell,
  row: number,
  col: number,
  date1904: boolean
): CellSnapshot | null {
  const model = cell.model;
  const formula = cell.formula; // triggers shared formula translation for slaves

  if (formula == null) {
    // Formula cell with no parseable formula — capture the cached result
    const cachedResult = convertFormulaResult(cell.result, date1904);
    return {
      row,
      col,
      value: cachedResult,
      formulaKind: "none",
      cachedResult
    };
  }

  // Determine formula kind
  const kind = classifyFormulaKind(model);

  // Capture the cached result from the XLSX
  const cachedResult = convertFormulaResult(cell.result, date1904);

  return {
    row,
    col,
    value: cachedResult,
    formulaKind: kind,
    formula,
    ref: model.ref,
    isDynamicArray: model.isDynamicArray ?? undefined,
    cachedResult
  };
}

/**
 * Classify a formula cell's kind based on its model properties.
 */
function classifyFormulaKind(model: {
  shareType?: string;
  ref?: string;
  formula?: string;
  sharedFormula?: string;
  isDynamicArray?: boolean;
}): FormulaCellKind {
  if (model.isDynamicArray) {
    return "dynamic-array";
  }

  if (model.shareType === "array" && model.ref) {
    return "cse";
  }

  if (model.shareType === "shared") {
    // shared-master has formula + ref, shared-slave has sharedFormula
    if (model.formula && model.ref) {
      return "shared-master";
    }
    if (model.sharedFormula) {
      return "shared-slave";
    }
  }

  return "normal";
}

// ============================================================================
// Value Conversion
// ============================================================================

/**
 * Convert a live cell value to a snapshot value.
 * - Dates → Excel serial number
 * - Rich text → plain string
 * - Errors → SnapshotErrorValue
 * - All other types pass through
 */
function convertCellValue(value: unknown, date1904: boolean): SnapshotCellValue {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "number") {
    return value;
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "boolean") {
    return value;
  }
  if (value instanceof Date) {
    return dateToExcel(value, date1904);
  }
  if (typeof value === "object" && "error" in value) {
    // Gate the error code through the known set so user-supplied values
    // like `{ error: "anything" }` can't pollute the snapshot. Unknown
    // strings fall back to #VALUE! — matches Excel's default diagnostic
    // when it reads an unrecognised error value from a persisted file.
    const raw = (value as { error: string }).error;
    const KNOWN_ERRORS = new Set([
      "#N/A",
      "#NULL!",
      "#DIV/0!",
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#NUM!",
      "#GETTING_DATA",
      "#CALC!",
      "#SPILL!",
      "#CONNECT!",
      "#BLOCKED!",
      "#UNKNOWN!",
      "#FIELD!",
      "#BUSY!"
    ]);
    const code = KNOWN_ERRORS.has(raw) ? raw : "#VALUE!";
    return { error: code } as SnapshotErrorValue;
  }
  // Rich text → plain string
  if (typeof value === "object" && "richText" in value) {
    return ((value as { richText: { text: string }[] }).richText || []).map(r => r.text).join("");
  }
  // Hyperlink / other objects with text
  if (typeof value === "object" && "text" in value) {
    return (value as { text: string }).text;
  }

  return null;
}

/**
 * Convert a formula result to a snapshot value.
 */
function convertFormulaResult(result: unknown, date1904: boolean): SnapshotCellValue {
  if (result === undefined || result === null) {
    return null;
  }
  return convertCellValue(result, date1904);
}

// ============================================================================
// Build Tables
// ============================================================================

function buildTables(ws: Worksheet): TableSnapshot[] {
  if (!ws.getTables) {
    return [];
  }

  const tables: TableSnapshot[] = [];

  for (const t of ws.getTables()) {
    const model = t.table;
    if (!model || !model.tl) {
      continue;
    }

    const columns: TableColumnSnapshot[] = (model.columns || []).map(c => ({
      name: c.name
    }));

    tables.push({
      name: model.name || model.displayName || "",
      columns,
      topLeft: { row: model.tl.row, col: model.tl.col },
      dataRowCount: (model.rows || []).length,
      hasHeaderRow: model.headerRow !== false,
      hasTotalsRow: model.totalsRow === true
    });
  }

  return tables;
}

// ============================================================================
// Build Defined Names
// ============================================================================

function buildDefinedNames(workbook: WorkbookLike): ReadonlyMap<string, DefinedNameSnapshot> {
  const map = new Map<string, DefinedNameSnapshot>();

  if (!workbook.definedNames) {
    return map;
  }

  // Build a sheet-id-to-name lookup for resolving localSheetId → sheet name.
  //
  // `localSheetId` stores the 0-based position in the XLSX-model sheet list,
  // which in this library corresponds to the index in `workbook.worksheets`
  // (the filtered+ordered view, see `workbook.xlsx/xform/book/workbook-xform.ts`
  // line 50-94 where `index` is incremented per rendered sheet).
  //
  // Iterate the array exactly once and capture the position directly — avoids
  // an O(n²) `indexOf` over a fresh array each call. Any caller that has
  // deleted sheets after creating scoped names is responsible for updating
  // `localSheetId` to match the new positions; this layer just reflects the
  // workbook's current state.
  const sheetIdToName = new Map<number, string>();
  const liveSheets = workbook.worksheets;
  for (let idx = 0; idx < liveSheets.length; idx++) {
    sheetIdToName.set(idx, liveSheets[idx].name);
  }

  // getAllEntries() returns self-contained entries — no second lookup needed.
  const entries = workbook.definedNames.getAllEntries();
  for (const entry of entries) {
    if (!entry.ranges || entry.ranges.length === 0) {
      continue;
    }

    // Convert numeric localSheetId → sheet name string for the snapshot
    let scope: string | undefined;
    if (entry.localSheetId !== undefined) {
      scope = sheetIdToName.get(entry.localSheetId);
    }

    const snapshot: DefinedNameSnapshot = {
      name: entry.name,
      ranges: [...entry.ranges],
      ...(scope ? { scope } : {})
    };

    const key = scope ? scopedNameKey(scope, entry.name) : entry.name.toUpperCase();
    if (!map.has(key)) {
      map.set(key, snapshot);
    }
  }

  return map;
}
