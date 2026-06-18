/**
 * `Pivot` namespace surface.
 *
 * `import { Pivot } from "@cj-tech-master/excelts/excel"` → `Pivot.add(ws, model)`.
 *
 * Also re-exports the public pivot-table model types so consumers reference
 * them through the namespace (`Pivot.Value`, `Pivot.Model`, …) instead of an
 * internal module path.
 */
export { addPivotTable as add } from "@excel/worksheet";

export type {
  PivotTable as Handle,
  PivotTableSource as Source,
  PivotTableValue as Value,
  PivotTableModel as Model,
  PivotTableSubtotal as Subtotal,
  PivotTableChartFormat as ChartFormat,
  DataField,
  CacheField,
  SharedItemValue,
  ParsedCacheDefinition,
  RecordValue
} from "@excel/pivot-table";
