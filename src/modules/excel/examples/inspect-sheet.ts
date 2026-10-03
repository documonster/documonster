/**
 * Describe the shape of every sheet in a workbook: its size, and for each data region its header row, its
 * columns, the first data value of each column and where its formulas are.
 *
 * This is a recipe, not a library function, because the interesting decision — which row is the header —
 * belongs to the application. The file often records it, though, so a sheet is split into **regions**, each
 * described on its own, and every region says where its header came from:
 *
 * | Region         | Header row                                 | `headerSource`                            |
 * | -------------- | ------------------------------------------ | ----------------------------------------- |
 * | A table        | declared by the table, totals row excluded | `"table"`                                 |
 * | An auto-filter | its first row — where the drop-downs sit   | `"autoFilter"`                            |
 * | The rest       | the caller's `headerRow`, else a guess     | `"explicit"`, `"frozenRows"`, `"default"` |
 *
 * A declaration in the file beats the caller, and the caller beats a guess: frozen rows only *keep* a row
 * in view, so the last frozen row is used as the header only when nobody said otherwise.
 *
 * Three things matter if you adapt it:
 *
 * - **Iterate when you scan.** `Worksheet.eachRow` and `Row.eachCell` visit only cells that exist and hand
 *   over the cell itself, so a sparse sheet is not walked as its bounding rectangle. Reading one known
 *   address — the first data row under a header — is fine with `Cell.getValue`: value readers create
 *   nothing.
 * - **Ask the raw value how a formula is held.** It says whether a cell defines its formula (`formula`), is
 *   a clone of a shared formula (`sharedFormula`; `Cell.getFormula` returns it translated to that cell) or
 *   anchors an array formula (`shareType: "array"`, with its `ref`). Only the anchor carries an array
 *   formula: the other cells of its `ref` read back as plain values, so they are attributed by position.
 * - **A value and its display are different answers.** `firstDataValue` is typed — a number stays a number,
 *   a formula is its cached result — and `firstDataText` is what Excel shows after the number format.
 *
 * Run without arguments, it builds a sample covering each of these cases and checks its own output.
 *
 * Usage (the example runner passes no arguments, so run it directly to pass any):
 *   node --import @oxc-node/core/register src/modules/excel/examples/inspect-sheet.ts [input.xlsx] [headerRow]
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { buildSample, verifySample } from "@excel/examples/utils/inspect-sheet-sample";
import { Address, Cell, Row, Table, Workbook, Worksheet } from "@excel/index";
import type { AutoFilter, CellErrorValue, CellValue } from "@excel/index";

export type HeaderSource = "table" | "autoFilter" | "explicit" | "frozenRows" | "default";
export type PlainValue = string | number | boolean | Date | CellErrorValue | null;
type FormulaKind = "own" | "shared" | "array";

export interface ColumnInfo {
  letter: string;
  header: string;
  /** The first data row's value, typed — see {@link plainValue}. `null` when that cell is empty. */
  firstDataValue: PlainValue;
  /** The same cell as Excel displays it. */
  firstDataText: string;
  /** The first non-empty value in the data rows, for a column whose first data row is blank. */
  sample: string;
  formulas: {
    /** Cells that carry their own formula. */
    own: number;
    /** Clones of a shared formula defined in another cell. */
    shared: number;
    /**
     * Cells of an array formula's range that exist on the sheet, the anchor included. Cells of the range the
     * file left out are not counted — reading them would mean inventing them.
     */
    array: number;
  };
  firstFormula?: { address: string; kind: FormulaKind; text: string; anchor?: string };
  /** A table's totals-row cell, as displayed. */
  total?: string;
}

export interface RegionInfo {
  kind: "table" | "autoFilter" | "sheet";
  name?: string;
  ref?: string;
  headerRow: number | null;
  headerSource: HeaderSource;
  firstDataRow: number;
  /** Last data row, or `null` for an open-ended region that runs to the end of the sheet. */
  lastDataRow: number | null;
  totalsRow: number | null;
  columns: ColumnInfo[];
}

export interface SheetInfo {
  sheet: string;
  /** The last row number in use. */
  rowCount: number;
  /** How many rows hold a value — fewer than `rowCount` on a sparse sheet. */
  actualRowCount: number;
  frozenRows: number;
  regions: RegionInfo[];
}

interface Box {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/** A region while it is being filled, its columns keyed by number. */
type WorkingRegion = Omit<RegionInfo, "columns"> & {
  columns: Map<number, ColumnInfo>;
  hasData: boolean;
};

/** A region the file declares, which therefore has bounds. */
type DeclaredRegion = WorkingRegion & {
  bounds: Box;
  /** A table's column names, left to right — declared, so reported even without cells. */
  declaredColumns?: string[];
};

interface ArrayFormula extends Box {
  anchor: string;
  text: string;
}

function box(ref: string): Box {
  const { s, e } = Address.decodeRange(ref);
  return { top: s.r + 1, bottom: e.r + 1, left: s.c + 1, right: e.c + 1 };
}

function contains(b: Box, r: number, c: number): boolean {
  return b.top <= r && r <= b.bottom && b.left <= c && c <= b.right;
}

function autoFilterRef(filter: AutoFilter | null): string | undefined {
  if (filter === null || typeof filter === "string") {
    return filter ?? undefined;
  }
  const corner = (c: string | { row: number; col: number }) =>
    typeof c === "string" ? c : Address.encodeCell({ r: c.row - 1, c: c.col - 1 });
  return `${corner(filter.from)}:${corner(filter.to)}`;
}

/**
 * A cell's value with its type kept, before number formatting:
 *
 * - numbers, booleans, strings and dates as themselves; a checkbox as its boolean;
 * - a formula as its cached result, `null` when the file cached none;
 * - an error as the `{ error }` value, whether stored or a formula's result — not its text, which would
 *   be indistinguishable from a string that happens to read `#N/A`;
 * - rich text and a hyperlink as their text;
 * - a merged cell as its master's value, which is what `Cell.getValue` reads for it.
 */
export function plainValue(ws: Worksheet.Handle, r: number, c: number): PlainValue {
  const value = Cell.getValue(ws, r, c);
  if (value === null || typeof value !== "object" || value instanceof Date) {
    return value ?? null;
  }
  if ("formula" in value || "sharedFormula" in value) {
    return value.result ?? null;
  }
  if ("error" in value) {
    return value;
  }
  if ("checkbox" in value) {
    return value.checkbox;
  }
  if ("richText" in value || "text" in value) {
    return Cell.getText(ws, r, c);
  }
  return null;
}

function isArrayAnchor(value: CellValue): value is CellValue & { formula: string; ref: string } {
  return (
    value !== null &&
    typeof value === "object" &&
    "shareType" in value &&
    value.shareType === "array"
  );
}

/**
 * Every array formula on the sheet, collected before anything is classified: the anchor may sit in a row
 * the description skips (a header, a totals row, a title) while the cells it covers are data.
 */
function arrayFormulas(ws: Worksheet.Handle): ArrayFormula[] {
  const found: ArrayFormula[] = [];
  Worksheet.eachRow(ws, (_row, r) => {
    Row.eachCell(ws, r, cell => {
      const value = Cell.view(cell).value;
      if (isArrayAnchor(value)) {
        found.push({ ...box(value.ref), anchor: cell.address, text: value.formula });
      }
    });
  });
  return found;
}

/** How the cell at `address` holds a formula, or `undefined` for a cell that holds none. */
function formulaOf(
  ws: Worksheet.Handle,
  value: CellValue,
  address: string,
  arrays: readonly ArrayFormula[],
  r: number,
  c: number
): ColumnInfo["firstFormula"] {
  if (value !== null && typeof value === "object" && !isArrayAnchor(value)) {
    if ("sharedFormula" in value) {
      return { address, kind: "shared", text: Cell.getFormula(ws, r, c) ?? "" };
    }
    if ("formula" in value) {
      return { address, kind: "own", text: value.formula };
    }
  }
  const array = arrays.find(a => contains(a, r, c));
  return array && { address, kind: "array", text: array.text, anchor: array.anchor };
}

export function describeSheet(ws: Worksheet.Handle, explicitHeaderRow?: number): SheetInfo {
  const arrays = arrayFormulas(ws);

  // The regions the file declares. Tables first, so a cell of a table inside a filtered range is the table's.
  const declared: DeclaredRegion[] = [];
  for (const t of Table.list(ws)) {
    const bounds = box(Table.ref(t));
    const hasHeader = Table.headerRow(t) !== false;
    const hasTotals = Table.totalsRow(t) === true;
    declared.push({
      kind: "table",
      name: Table.name(t),
      ref: Table.ref(t),
      headerRow: hasHeader ? bounds.top : null,
      headerSource: "table",
      firstDataRow: hasHeader ? bounds.top + 1 : bounds.top,
      lastDataRow: hasTotals ? bounds.bottom - 1 : bounds.bottom,
      totalsRow: hasTotals ? bounds.bottom : null,
      bounds,
      columns: new Map(),
      declaredColumns: Table.model(t).columns.map(column => column.name),
      hasData: false
    });
  }
  const filterRef = autoFilterRef(Worksheet.autoFilter(ws));
  if (filterRef) {
    const bounds = box(filterRef);
    declared.push({
      kind: "autoFilter",
      ref: filterRef,
      headerRow: bounds.top,
      headerSource: "autoFilter",
      firstDataRow: bounds.top + 1,
      lastDataRow: bounds.bottom,
      totalsRow: null,
      bounds,
      columns: new Map(),
      hasData: false
    });
  }

  // Everything else. Its header is the caller's word, else a guess.
  const panes = Worksheet.panes(ws);
  const frozenRows = panes?.state === "frozen" ? (panes.ySplit ?? 0) : 0;
  const [headerRow, headerSource]: [number, HeaderSource] =
    explicitHeaderRow !== undefined
      ? [explicitHeaderRow, "explicit"]
      : frozenRows > 0
        ? [frozenRows, "frozenRows"]
        : [1, "default"];
  const rest: WorkingRegion = {
    kind: "sheet",
    headerRow,
    headerSource,
    firstDataRow: headerRow + 1,
    lastDataRow: null,
    totalsRow: null,
    columns: new Map(),
    hasData: false
  };

  const regionAt = (r: number, c: number): WorkingRegion =>
    declared.find(region => contains(region.bounds, r, c)) ?? rest;

  function columnOf(region: WorkingRegion, c: number): ColumnInfo {
    let info = region.columns.get(c);
    if (!info) {
      const first = region.firstDataRow;
      // The cell under the header may belong to another region — a table below the remainder's header,
      // say — and then it is not this region's first data value.
      const own =
        (region.lastDataRow === null || first <= region.lastDataRow) &&
        regionAt(first, c) === region;
      info = {
        letter: Address.encodeCol(c - 1),
        header: "",
        // One known address, read directly: nothing is created if the cell is empty.
        firstDataValue: own ? plainValue(ws, first, c) : null,
        firstDataText: own ? Cell.getDisplayText(ws, first, c) : "",
        sample: "",
        formulas: { own: 0, shared: 0, array: 0 }
      };
      region.columns.set(c, info);
    }
    return info;
  }

  // A table's columns are declared, so they are reported even when none of their cells exists — a table
  // without a header row, or without data, still has them. A header cell, when there is one, is read below.
  for (const region of declared) {
    region.declaredColumns?.forEach((name, j) => {
      columnOf(region, region.bounds.left + j).header = name;
    });
  }

  Worksheet.eachRow(ws, (_row, r) => {
    Row.eachCell(ws, r, (cell, c) => {
      const region = regionAt(r, c);
      if (region.headerRow !== null && r < region.headerRow) {
        return; // a title or note above the header, not part of the region
      }
      const info = columnOf(region, c);
      if (r === region.headerRow) {
        info.header = Cell.view(cell).text;
        return;
      }
      if (r === region.totalsRow) {
        info.total = Cell.getDisplayText(ws, r, c);
        return;
      }
      region.hasData = true;
      info.sample ||= Cell.getDisplayText(ws, r, c);

      const formula = formulaOf(ws, Cell.view(cell).value, cell.address, arrays, r, c);
      if (formula) {
        info.formulas[formula.kind]++;
        info.firstFormula ??= formula;
      }
    });
  });

  // A declared region is reported even when empty — that is a fact about the file. The remainder is
  // reported only when it holds data below its header; otherwise its "header" is just a title.
  const reported: WorkingRegion[] = rest.hasData ? [...declared, rest] : declared;
  return {
    sheet: Worksheet.getName(ws),
    rowCount: Worksheet.rowCount(ws),
    actualRowCount: Worksheet.actualRowCount(ws),
    frozenRows,
    regions: reported.map(region => ({
      kind: region.kind,
      ...(region.name === undefined ? {} : { name: region.name }),
      ...(region.ref === undefined ? {} : { ref: region.ref }),
      headerRow: region.headerRow,
      headerSource: region.headerSource,
      firstDataRow: region.firstDataRow,
      lastDataRow: region.lastDataRow,
      totalsRow: region.totalsRow,
      columns: [...region.columns.entries()].sort(([a], [b]) => a - b).map(([, info]) => info)
    }))
  };
}

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/excel-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const input = process.argv[2] ?? path.join(outDir, "inspect-sheet.xlsx");
const headerArg = process.argv[3];
const explicitHeaderRow = headerArg === undefined ? undefined : Number(headerArg);
if (
  explicitHeaderRow !== undefined &&
  (!Number.isInteger(explicitHeaderRow) || explicitHeaderRow < 1)
) {
  throw new RangeError(`headerRow must be a positive integer, got ${headerArg}`);
}
if (!process.argv[2]) {
  await buildSample(input);
}

const wb = Workbook.create();
await Workbook.readFile(wb, input);
const report = Workbook.getWorksheets(wb).map(ws => describeSheet(ws, explicitHeaderRow));
console.log(JSON.stringify({ workbook: path.basename(input), sheets: report }, null, 2));
if (!process.argv[2]) {
  verifySample(report);
}
