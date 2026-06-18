/**
 * Excel → formula `WorkbookLike` adapter.
 *
 * The formula engine consumes the structural `WorkbookLike` / `WorksheetLike` /
 * `CellLike` contract (`@formula`). documonster' workbook/worksheet/cell are
 * plain-data records with no getters or methods, so this module wraps a
 * `WorkbookData` into a live `WorkbookLike` view. Reads delegate to the flat
 * container/cell helpers; cell `value` / `result` setters write back to the
 * real `CellData` so `calculateFormulas` mutates the workbook in place.
 *
 * Excel (layer 4) may import formula (layer 3); this is the sanctioned seam.
 */
import {
  type CellData,
  cellFormula,
  cellGetValue,
  cellResult,
  cellSetResult,
  cellSetValue,
  cellType
} from "@excel/cell";
import { definedNamesGetAllEntries, definedNamesGetAllNames } from "@excel/defined-names";
import { type RowData } from "@excel/row";
import { type WorkbookData, getWorksheets } from "@excel/workbook.browser";
import {
  type WorksheetData,
  eachRow,
  findCell,
  getCell,
  getSheetDimensions,
  getTables
} from "@excel/worksheet";
import { calculateFormulas as calculateFormulasEngine } from "@formula/integration/calculate-formulas";
import type {
  CellLike,
  RowLike,
  WorkbookLike,
  WorksheetLike,
  DefinedNamesLike
} from "@formula/materialize/types";

function wrapCell(c: CellData): CellLike {
  return {
    get row() {
      return c.row.number;
    },
    get col() {
      return c.column.number;
    },
    get type() {
      return cellType(c) as unknown as CellLike["type"];
    },
    get formula() {
      return cellFormula(c);
    },
    get model() {
      return c._value.model as unknown as CellLike["model"];
    },
    get value() {
      return cellGetValue(c);
    },
    set value(v: unknown) {
      cellSetValue(c, v as never);
    },
    get result() {
      return cellResult(c) as unknown as CellLike["result"];
    },
    set result(v: CellLike["result"]) {
      cellSetResult(c, v as never);
    }
  };
}

function wrapRow(r: RowData): RowLike {
  return {
    hidden: r.hidden,
    get cells(): readonly (CellLike | undefined)[] {
      return r.cells.map(c => (c ? wrapCell(c) : undefined));
    }
  };
}

function wrapWorksheet(ws: WorksheetData): WorksheetLike {
  const dims = getSheetDimensions(ws);
  const hasDim = dims.top <= dims.bottom && dims.left <= dims.right;
  return {
    id: ws.id,
    name: ws._name,
    dimensions: hasDim
      ? { top: dims.top, left: dims.left, bottom: dims.bottom, right: dims.right }
      : null,
    eachRow(
      optOrCb: { includeEmpty?: boolean } | ((row: RowLike, rowNumber: number) => void),
      maybeCb?: (row: RowLike, rowNumber: number) => void
    ): void {
      const opts = typeof optOrCb === "function" ? {} : optOrCb;
      const cb = (typeof optOrCb === "function" ? optOrCb : maybeCb)!;
      eachRow(ws, opts, (row, n) => cb(wrapRow(row), n));
    },
    findCell(row: number, col: number): CellLike | undefined {
      const c = findCell(ws, row, col);
      return c ? wrapCell(c) : undefined;
    },
    getCell(row: number, col: number): CellLike {
      return wrapCell(getCell(ws, row, col));
    },
    getTables() {
      return getTables(ws) as unknown as ReturnType<NonNullable<WorksheetLike["getTables"]>>;
    },
    get mergedRegions(): readonly { top: number; left: number; bottom: number; right: number }[] {
      return Object.values(ws._merges).map(m => ({
        top: m.top,
        left: m.left,
        bottom: m.bottom,
        right: m.right
      }));
    }
  };
}

/** Wrap an documonster `WorkbookData` record as a formula-engine `WorkbookLike`.
 *
 * Memoised per `WorkbookData` so repeated calls return the SAME `WorkbookLike`
 * instance. The formula engine persists spill state in a `WeakMap` keyed by the
 * workbook object; a stable adapter identity across `calculateFormulas` calls is
 * required for spill-ghost cleanup to work. */
const adapterCache = new WeakMap<WorkbookData, WorkbookLike>();

export function toWorkbookLike(wb: WorkbookData): WorkbookLike {
  const cached = adapterCache.get(wb);
  if (cached) {
    return cached;
  }
  const live = buildWorkbookLike(wb);
  adapterCache.set(wb, live);
  return live;
}

function buildWorkbookLike(wb: WorkbookData): WorkbookLike {
  return {
    get worksheets(): WorksheetLike[] {
      return getWorksheets(wb).map(wrapWorksheet);
    },
    getWorksheet(id?: number | string): WorksheetLike | undefined {
      const wsLikes = getWorksheets(wb).map(wrapWorksheet);
      if (id === undefined) {
        return wsLikes[0];
      }
      if (typeof id === "number") {
        return wsLikes.find(w => w.id === id);
      }
      const lower = id.toLowerCase();
      return wsLikes.find(w => w.name.toLowerCase() === lower);
    },
    get definedNames(): DefinedNamesLike | undefined {
      return wb._definedNames
        ? {
            getAllEntries: () => definedNamesGetAllEntries(wb._definedNames) as never,
            getAllNames: () => definedNamesGetAllNames(wb._definedNames) as never
          }
        : undefined;
    },
    get calcProperties() {
      return wb.calcProperties as WorkbookLike["calcProperties"];
    },
    get properties() {
      return wb.properties as WorkbookLike["properties"];
    },
    get userFunctions() {
      return wb.userFunctions as WorkbookLike["userFunctions"];
    }
  };
}

/**
 * Recalculate all formulas in a workbook, mutating cached results in place.
 * Excel-side wrapper around the formula engine that adapts the plain-data
 * `WorkbookData` to the engine's `WorkbookLike` contract.
 */
export function calculateFormulas(wb: WorkbookData): void {
  calculateFormulasEngine(toWorkbookLike(wb));
}
