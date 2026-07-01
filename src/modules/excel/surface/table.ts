/**
 * `Table` namespace surface.
 *
 * Worksheet-level management:
 *   `Table.add(ws, model)`, `Table.get(ws, name)`, `Table.list(ws)`,
 *   `Table.remove(ws, name)`.
 * Table-handle operations:
 *   `Table.create(...)`, `Table.addRow(t, row)`, `Table.setName(t, name)`,
 *   `Table.model(t)`, `Table.column(t, idx)`, …
 */
export {
  addTable as add,
  getTable as get,
  getTables as list,
  removeTable as remove
} from "@excel/core/worksheet";

export {
  createTable as create,
  tableModel as model,
  tableSetModel as setModel,
  tableCommit as commit,
  tableAddRow as addRow,
  tableRemoveRows as removeRows,
  tableGetColumn as column,
  tableAddColumn as addColumn,
  tableRemoveColumns as removeColumns,
  tableRef as ref,
  tableSetRef as setRef,
  tableName as name,
  tableSetName as setName,
  tableDisplayName as displayName,
  tableSetDisplayName as setDisplayName,
  tableHeaderRow as headerRow,
  tableSetHeaderRow as setHeaderRow,
  tableTotalsRow as totalsRow,
  tableSetTotalsRow as setTotalsRow,
  tableTheme as theme,
  tableSetTheme as setTheme,
  tableShowFirstColumn as showFirstColumn,
  tableSetShowFirstColumn as setShowFirstColumn,
  tableShowLastColumn as showLastColumn,
  tableSetShowLastColumn as setShowLastColumn,
  tableShowRowStripes as showRowStripes,
  tableSetShowRowStripes as setShowRowStripes,
  tableShowColumnStripes as showColumnStripes,
  tableSetShowColumnStripes as setShowColumnStripes,
  tableColumnName as columnName,
  tableColumnSetName as columnSetName,
  tableColumnFilterButton as columnFilterButton,
  tableColumnSetFilterButton as columnSetFilterButton,
  tableColumnStyle as columnStyle,
  tableColumnSetStyle as columnSetStyle,
  tableColumnTotalsRowLabel as columnTotalsRowLabel,
  tableColumnSetTotalsRowLabel as columnSetTotalsRowLabel,
  tableColumnTotalsRowFunction as columnTotalsRowFunction,
  tableColumnSetTotalsRowFunction as columnSetTotalsRowFunction,
  tableColumnTotalsRowResult as columnTotalsRowResult,
  tableColumnSetTotalsRowResult as columnSetTotalsRowResult,
  tableColumnTotalsRowFormula as columnTotalsRowFormula,
  tableColumnSetTotalsRowFormula as columnSetTotalsRowFormula
} from "@excel/core/table";

/** A table handle. */
export type { TableData as Handle } from "@excel/core/table";
