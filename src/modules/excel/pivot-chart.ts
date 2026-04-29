import type {
  AddChartOptions,
  AddComboChartOptions,
  PivotChartOptions as PivotChartMetadataOptions,
  PivotChartSource
} from "@excel/chart/types";
import type { PivotTable } from "@excel/pivot-table";
import { quoteSheetName } from "@excel/utils/address";

const DEFAULT_PIVOT_CHART_PIVOT_AREA_XML =
  '<pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea>';

type PivotChartCreationOptions = AddChartOptions | AddComboChartOptions;

export function withPivotChartSource<T extends PivotChartCreationOptions>(
  pivotTable: PivotTable,
  options: T,
  fmtId?: number,
  pivotChartOptions?: PivotChartMetadataOptions
): T & { pivotSource: PivotChartSource } {
  // Auto-assign the next available fmtId when the caller didn't pass
  // one. Before this fix, a second `addPivotChart(pivotTable, …)`
  // against the same pivot would reuse `fmtId=0`, collide with the
  // first chart's `chartFormat` entry (the `exists` guard in
  // `ensurePivotChartFormat` would no-op), and produce two charts
  // pointing at a single pivotArea declaration — Excel renders only
  // one of them correctly.
  const existingFormats = pivotTable.chartFormats ?? [];
  const allocatedFmtId =
    fmtId ??
    (existingFormats.length === 0 ? 0 : Math.max(...existingFormats.map(f => f.format ?? 0)) + 1);
  ensurePivotChartFormat(pivotTable, allocatedFmtId);
  applyPivotChartOptions(pivotTable, pivotChartOptions ?? options.pivotChartOptions);
  return {
    ...options,
    pivotSource: {
      name: buildPivotChartSourceName(pivotTable),
      fmtId: allocatedFmtId,
      options: pivotChartOptions ?? options.pivotChartOptions
    }
  };
}

function applyPivotChartOptions(
  pivotTable: PivotTable,
  options: PivotChartMetadataOptions | undefined
): void {
  if (!options) {
    return;
  }
  pivotTable.pivotChartOptions = options;
  if (options.refreshOnOpen && pivotTable.cacheDefinition) {
    pivotTable.cacheDefinition.refreshOnLoad = "1";
  }
}

function ensurePivotChartFormat(pivotTable: PivotTable, fmtId: number): void {
  pivotTable.chartFormat ??= 1;
  pivotTable.chartFormats ??= [];
  const exists = pivotTable.chartFormats.some(
    format => format.chart === 0 && format.format === fmtId
  );
  if (!exists) {
    pivotTable.chartFormats.push({
      chart: 0,
      format: fmtId,
      series: true,
      pivotAreaXml: DEFAULT_PIVOT_CHART_PIVOT_AREA_XML
    });
  }
}

function buildPivotChartSourceName(pivotTable: PivotTable): string {
  const tableName = pivotTable.name ?? `PivotTable${pivotTable.tableNumber}`;
  return pivotTable.worksheetName
    ? `${quoteSheetName(pivotTable.worksheetName)}!${tableName}`
    : tableName;
}
