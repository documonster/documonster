import type {
  AddChartOptions,
  AddComboChartOptions,
  PivotChartOptions as PivotChartMetadataOptions,
  PivotChartSource
} from "@excel/chart/types";
import type { PivotTable } from "@excel/pivot-table";

const DEFAULT_PIVOT_CHART_PIVOT_AREA_XML =
  '<pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea>';

type PivotChartCreationOptions = AddChartOptions | AddComboChartOptions;

export function withPivotChartSource<T extends PivotChartCreationOptions>(
  pivotTable: PivotTable,
  options: T,
  fmtId = 0,
  pivotChartOptions?: PivotChartMetadataOptions
): T & { pivotSource: PivotChartSource } {
  ensurePivotChartFormat(pivotTable, fmtId);
  applyPivotChartOptions(pivotTable, pivotChartOptions ?? options.pivotChartOptions);
  return {
    ...options,
    pivotSource: {
      name: buildPivotChartSourceName(pivotTable),
      fmtId,
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

function quoteSheetName(sheetName: string): string {
  return /^[a-zA-Z0-9]+$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}
