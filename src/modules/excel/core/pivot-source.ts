/**
 * Which pivot tables read their data from a sheet or a table — so the sheet or the table is not removed from under
 * them. A pivot cache names its source (`<worksheetSource sheet="Data" ref="A1:C9"/>` or `name="Sales"`); with the
 * source gone the file still opened, but the pivot table could never be refreshed again.
 *
 * Its own file so that removing a sheet or a table does not pull the pivot-table builder into a bundle.
 */
import type { PivotTable } from "@excel/core/pivot-table";
import type { TableData } from "@excel/core/table";
import type { WorkbookData } from "@excel/core/workbook-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { ExcelPivotTableError } from "@excel/errors";
import { colCache } from "@excel/utils/col-cache";
import { excelNameKey, sheetNameKey } from "@utils/name-key";

function sourceSheetOf(pivot: PivotTable): string | undefined {
  return pivot.source?.name ?? pivot.cacheDefinition?.sourceSheet;
}

function sourceTableOf(pivot: PivotTable): string | undefined {
  return pivot.source?.tableName ?? pivot.cacheDefinition?.sourceTableName;
}

function pivotName(pivot: PivotTable): string {
  return pivot.name ?? `PivotTable${pivot.tableNumber}`;
}

function tableNames(table: TableData): Set<string> {
  const names = new Set([excelNameKey(table.table.name)]);
  if (table.table.displayName) {
    names.add(excelNameKey(table.table.displayName));
  }
  return names;
}

/**
 * Refuse to remove `target.table` — or, with no table, the sheet `target.sheet` — while a pivot table reads from it.
 * A pivot table on the sheet being removed goes with it, so it does not hold the sheet.
 *
 * @throws {ExcelPivotTableError} naming the pivot table and what it reads.
 */
export function checkNotPivotSource(
  wb: WorkbookData,
  target: { sheet: WorksheetData; table?: TableData },
  where: string
): void {
  const pivots = wb.pivotTables;
  if (!pivots || pivots.length === 0) {
    return;
  }
  const sheetKey = sheetNameKey(target.sheet._name);
  const tables = target.table ? [target.table] : Object.values(target.sheet.tables ?? {});
  const tableKeys = new Set(tables.flatMap(table => [...tableNames(table)]));
  for (const pivot of pivots) {
    if (!target.table && pivot.worksheetName !== undefined) {
      if (sheetNameKey(pivot.worksheetName) === sheetKey) {
        continue;
      }
    }
    const sourceTable = sourceTableOf(pivot);
    if (sourceTable !== undefined && tableKeys.has(excelNameKey(sourceTable))) {
      throw new ExcelPivotTableError(
        `${where}: pivot table "${pivotName(pivot)}" reads its data from table "${sourceTable}"; remove the pivot table first (Pivot.remove).`
      );
    }
    const sourceSheet = sourceSheetOf(pivot);
    if (
      !target.table &&
      sourceTable === undefined &&
      sourceSheet !== undefined &&
      sheetNameKey(sourceSheet) === sheetKey
    ) {
      throw new ExcelPivotTableError(
        `${where}: pivot table "${pivotName(pivot)}" reads its data from sheet "${target.sheet._name}"; remove the pivot table first (Pivot.remove).`
      );
    }
  }
}

interface Block {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * The cells a new pivot table occupies when it is written, before Excel refreshes it: its page filters (two columns,
 * then a blank row), the header row and one data row, across its row-label columns and data columns — the
 * placeholder `PivotTableXform` writes. A pivot table read from a file is left out; its extent is the file's.
 */
function placeholderBlock(pivot: PivotTable): Block | undefined {
  if (pivot.isLoaded) {
    return undefined;
  }
  const { row, col } = pivot.ref ? colCache.decodeAddress(pivot.ref) : { row: 3, col: 1 };
  const pages = pivot.pages?.length ?? 0;
  const pageOffset = pages > 0 ? pages + 1 : 0;
  const dataColumns =
    pivot.columns.length === 0 && pivot.values.length > 1 ? pivot.values.length : 1;
  const right = col + pivot.rows.length + dataColumns - 1;
  return {
    top: row,
    left: col,
    bottom: row + pageOffset + 1,
    right: pages > 0 ? Math.max(right, col + 1) : right
  };
}

/**
 * Refuse a new pivot table whose cells overlap another new pivot table on the same sheet. Excel refuses to refresh
 * two overlapping pivots ("There's already a PivotTable there"), and the most common way to get there is two pivots
 * left at the default anchor `A3`. A pivot that is refreshed grows from its anchor, so room has to be left beyond
 * this; only the overlap that is certain is refused.
 */
export function checkPivotPlacement(sheetPivots: readonly PivotTable[], pivot: PivotTable): void {
  const block = placeholderBlock(pivot);
  if (!block) {
    return;
  }
  for (const other of sheetPivots) {
    const o = placeholderBlock(other);
    if (
      o &&
      block.left <= o.right &&
      o.left <= block.right &&
      block.top <= o.bottom &&
      o.top <= block.bottom
    ) {
      throw new ExcelPivotTableError(
        `Pivot table "${pivotName(pivot)}" at ${pivot.ref ?? "A3"} overlaps pivot table "${pivotName(other)}" at ` +
          `${other.ref ?? "A3"}; give each pivot table on a sheet its own \`ref\`, with room to grow.`
      );
    }
  }
}
