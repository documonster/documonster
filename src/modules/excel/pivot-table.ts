import { PivotTableError } from "@excel/errors";
import type { Table } from "@excel/table";
import { colCache } from "@excel/utils/col-cache";
import { range, toSortedArray } from "@utils/utils";

/**
 * Interface representing the source data abstraction for pivot tables.
 * This allows both Worksheet and Table to be used as pivot table data sources.
 */
export interface PivotTableSource {
  /** Name of the worksheet containing the source data (used in pivotCacheDefinition) */
  name: string;
  /**
   * Name of the source Table (e.g., "SalesData").
   * When present, pivotCacheDefinition uses `<worksheetSource name="..."/>` instead of ref+sheet.
   */
  tableName?: string;
  /** Get row values by 1-indexed row number */
  getRow(rowNumber: number): { values: unknown[] };
  /** Get column values by 1-indexed column number */
  getColumn(columnNumber: number): { values: unknown[] };
  /** Get all sheet values as a sparse 2D array */
  getSheetValues(): unknown[][];
  /** Dimensions with short range reference (e.g., "A1:E10") */
  dimensions: { shortRange: string };
}

/**
 * A value field specification with optional per-field metric override.
 * Use this instead of a plain string when you need different aggregation
 * metrics for individual value fields.
 *
 * @example
 * ```ts
 * values: [
 *   { name: "Sales", metric: "sum" },
 *   { name: "Quantity", metric: "count" },
 *   { name: "Price", metric: "average" },
 * ]
 * ```
 */
export interface PivotTableValue {
  /** Column name to aggregate */
  name: string;
  /**
   * Aggregation metric for this specific value field.
   * Overrides the table-wide `metric` when specified.
   * @default inherited from PivotTableModel.metric (which defaults to 'sum')
   */
  metric?: PivotTableSubtotal;
}

/**
 * Model for creating a new pivot table.
 * Pass this to worksheet.addPivotTable() to create a pivot table.
 */
export interface PivotTableModel {
  /**
   * Source worksheet for the pivot table data.
   * Either sourceSheet or sourceTable must be provided (mutually exclusive).
   */
  sourceSheet?: PivotTableSource;
  /**
   * Source table for the pivot table data.
   * Either sourceSheet or sourceTable must be provided (mutually exclusive).
   * The table must have headerRow=true and contain at least one data row.
   */
  sourceTable?: Table;
  /** Column names to use as row fields in the pivot table */
  rows: string[];
  /**
   * Column names to use as column fields in the pivot table.
   * If omitted or empty, Excel will use "Values" as the column field.
   * When multiple values are specified alongside columns, the synthetic
   * "∑Values" pseudo-field is appended to the column axis automatically.
   * @default []
   */
  columns?: string[];
  /**
   * Column names (or value field specifications) to aggregate in the pivot table.
   * Each entry can be a plain string (column name) or a {@link PivotTableValue}
   * object with a per-field metric override.
   *
   * @example
   * ```ts
   * // Simple: all values use the table-wide metric
   * values: ["Sales", "Quantity"]
   *
   * // Per-value metrics
   * values: [
   *   { name: "Sales", metric: "sum" },
   *   { name: "Quantity", metric: "count" },
   *   { name: "Price", metric: "average" },
   * ]
   *
   * // Mixed: plain strings inherit table-wide metric
   * values: ["Sales", { name: "Quantity", metric: "count" }]
   * ```
   */
  values: (string | PivotTableValue)[];
  /**
   * Column names to use as page fields (report filters) in the pivot table.
   * Page fields appear as dropdown filters above the pivot table.
   * @default []
   */
  pages?: string[];
  /**
   * Default aggregation metric for all value fields.
   * Individual value fields can override this via {@link PivotTableValue.metric}.
   * @default 'sum'
   */
  metric?: PivotTableSubtotal;
  /**
   * Controls whether pivot table style overrides worksheet column widths.
   * - '0': Preserve worksheet column widths (useful for custom sizing)
   * - '1': Apply pivot table style width/height (default Excel behavior)
   * @default '1'
   */
  applyWidthHeightFormats?: "0" | "1";
}

/** Allowed element types within CacheField.sharedItems */
export type SharedItemValue = string | number | boolean | Date | PivotErrorValue | null;

/**
 * Wrapper for OOXML error values in sharedItems (e.g. `<e v="REF!"/>`).
 * Distinguishes error strings from regular strings so they roundtrip as `<e>` not `<s>`.
 */
export class PivotErrorValue {
  /** The error code without the leading '#' (e.g. "REF!", "VALUE!", "N/A") */
  readonly code: string;
  constructor(code: string) {
    this.code = code;
  }
  /** Returns the display form with '#' prefix, e.g. "#REF!" */
  toString(): string {
    return `#${this.code}`;
  }
}

/**
 * Represents a cache field in a pivot table.
 * Cache fields store unique values from source columns for row/column grouping.
 */
export interface CacheField {
  /** Name of the field (column header from source) */
  name: string;
  /** Unique values for row/column fields, null for value fields */
  sharedItems: SharedItemValue[] | null;
  /** Whether the field contains numeric values (raw attribute string for roundtrip: "0" or "1") */
  containsNumber?: string;
  /** Whether the field contains only integer values (raw attribute string for roundtrip: "0" or "1") */
  containsInteger?: string;
  /** Minimum value for numeric fields */
  minValue?: number;
  /** Maximum value for numeric fields */
  maxValue?: number;
  /** Number format ID (preserved on roundtrip, defaults to "0") */
  numFmtId?: string;
  // ----- Loaded sharedItems attribute preservation (roundtrip fidelity) -----
  /** Original containsSemiMixedTypes attribute from loaded file */
  containsSemiMixedTypes?: string;
  /** Original containsNonDate attribute from loaded file */
  containsNonDate?: string;
  /** Original containsString attribute from loaded file */
  containsString?: string;
  /** Original containsBlank attribute from loaded file */
  containsBlank?: string;
  /** Original containsDate attribute from loaded file */
  containsDate?: string;
  /** Original containsMixedTypes attribute from loaded file */
  containsMixedTypes?: string;
  /** Flag indicating this cache field was loaded from file */
  isLoaded?: boolean;
  /** Preserved <fieldGroup> raw XML for roundtrip (loaded models only) */
  fieldGroupXml?: string;
  /** Bag of additional cacheField attributes not individually modeled (for roundtrip preservation) */
  extraAttrs?: Record<string, string>;
}

/** Aggregation function types for pivot table data fields */
export type PivotTableSubtotal =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNums"
  | "stdDev"
  | "stdDevP"
  | "var"
  | "varP";

/** Map from PivotTableSubtotal to its Excel display name prefix */
export const METRIC_DISPLAY_NAMES: Readonly<Record<PivotTableSubtotal, string>> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
  product: "Product",
  countNums: "Count Numbers",
  stdDev: "StdDev",
  stdDevP: "StdDevP",
  var: "Var",
  varP: "VarP"
};

/** Set of all valid PivotTableSubtotal values (for runtime validation) */
export const VALID_SUBTOTALS: ReadonlySet<string> = new Set<string>(
  Object.keys(METRIC_DISPLAY_NAMES)
);

/**
 * Data field configuration for pivot table aggregation.
 * Defines how values are aggregated in the pivot table.
 */
export interface DataField {
  /** Display name for the data field (e.g., "Sum of Sales") */
  name: string;
  /** Index of the source field in cacheFields */
  fld: number;
  /** Base field index for calculated fields */
  baseField?: number;
  /** Base item index for calculated fields */
  baseItem?: number;
  /** Aggregation function (default: 'sum') */
  subtotal?: PivotTableSubtotal;
  /** Number format ID (preserved on roundtrip for currency/date formatting) */
  numFmtId?: number;
}

/**
 * Internal pivot table representation used by the library.
 * This is the processed model after calling makePivotTable().
 */
export interface PivotTable {
  /** Source data adapter (always present for new pivot tables) */
  source?: PivotTableSource;
  /** Field indices for row fields */
  rows: number[];
  /** Field indices for column fields */
  columns: number[];
  /** Field indices for value fields */
  values: number[];
  /** Field indices for page fields (report filters) */
  pages?: number[];
  /** Default aggregation metric */
  metric: PivotTableSubtotal;
  /** Per-value metric overrides (parallel to `values` array). Falls back to `metric`. */
  valueMetrics: PivotTableSubtotal[];
  /** Cache fields with shared items */
  cacheFields: CacheField[];
  /** Cache ID for linking to pivot cache */
  cacheId: string;
  /** Width/height format setting */
  applyWidthHeightFormats: "0" | "1";
  /** 1-indexed table number for file naming (pivotTable1.xml, pivotTable2.xml, etc.) */
  tableNumber: number;
  /** Workbook relationship ID, assigned during write by addWorkbookRels() */
  rId?: string;
  /** Flag indicating this pivot table was loaded from file (not newly created) */
  isLoaded?: boolean;
  /** Data fields for loaded pivot tables */
  dataFields?: DataField[];
  /** Cache definition for loaded pivot tables */
  cacheDefinition?: ParsedCacheDefinition;
  /** Cache records for loaded pivot tables */
  cacheRecords?: ParsedCacheRecords;
}

/**
 * Parsed cache definition from loaded pivot table files.
 */
export interface ParsedCacheDefinition {
  sourceRef?: string;
  sourceSheet?: string;
  /** Source table name (name style - references a named Table) */
  sourceTableName?: string;
  /** Cache source type (default "worksheet") */
  cacheSourceType?: string;
  cacheFields: CacheField[];
  recordCount?: number;
  rId?: string;
  /** Additional attributes to preserve */
  refreshOnLoad?: string;
  createdVersion?: string;
  refreshedVersion?: string;
  minRefreshableVersion?: string;
  isLoaded?: boolean;
  // ----- BUG-26: Additional root attributes -----
  backgroundQuery?: string;
  supportSubquery?: string;
  supportAdvancedDrill?: string;
  /** Bag of additional root attributes not individually modeled (for roundtrip) */
  extraRootAttrs?: Record<string, string>;
  // ----- BUG-28: worksheetSource extra attributes -----
  /** worksheetSource r:id attribute (for external connections) */
  worksheetSourceRId?: string;
  // ----- BUG-29: cache definition extLst raw XML -----
  extLstXml?: string;
  // ----- R6-BugA: catch-all unknown child elements raw XML -----
  /** Preserved unknown child elements XML for roundtrip (e.g. calculatedItems, cacheHierarchies) */
  unknownElementsXml?: string;
  // ----- R8-B9: non-worksheet cacheSource children raw XML -----
  /** Preserved raw XML for non-worksheetSource children inside <cacheSource> (e.g. <consolidation>) */
  cacheSourceXml?: string;
}

/** Allowed element types within cache record values */
export type RecordValue =
  | { type: "x"; value: number }
  | { type: "n"; value: number }
  | { type: "s"; value: string }
  | { type: "b"; value: boolean }
  | { type: "m" }
  | { type: "d"; value: Date }
  | { type: "e"; value: string };

/**
 * Parsed cache records from loaded pivot table files.
 */
export interface ParsedCacheRecords {
  records: RecordValue[][];
  count: number;
  isLoaded?: boolean;
  // R8-B11: Preserved original root element attributes for roundtrip
  /** Extra root attributes beyond xmlns/xmlns:r/count (for roundtrip preservation) */
  extraRootAttrs?: Record<string, string>;
}

/**
 * Data rows start at index 2 in ExcelJS sparse arrays:
 * index 0 = empty (ExcelJS convention), index 1 = header row.
 */
const DATA_START_INDEX = 2;

/**
 * Creates a PivotTableSource adapter from a Table object.
 * This allows Tables to be used as pivot table data sources with the same interface as Worksheets.
 */
function createTableSourceAdapter(table: Table): PivotTableSource {
  const tableModel = table.model;

  // Validate that table has headerRow enabled (required for pivot table column names)
  if (tableModel.headerRow === false) {
    throw new PivotTableError(
      "Cannot create pivot table from a table without headers. Set headerRow: true on the table."
    );
  }

  // Validate table has data rows
  if (!tableModel.rows || tableModel.rows.length === 0) {
    throw new PivotTableError(
      "Cannot create pivot table from an empty table. Add data rows to the table."
    );
  }

  const columnNames = tableModel.columns.map(col => col.name);

  // Check for duplicate column names
  const nameSet = new Set<string>();
  for (const name of columnNames) {
    if (nameSet.has(name)) {
      throw new PivotTableError(
        `Duplicate column name "${name}" found in table. Pivot tables require unique column names.`
      );
    }
    nameSet.add(name);
  }

  // Build the full data array: headers + rows
  const headerRow = [undefined, ...columnNames]; // sparse array starting at index 1
  const dataRows = tableModel.rows.map(row => [undefined, ...row]); // sparse array starting at index 1

  // Calculate the range reference for the table
  const tl = tableModel.tl;
  if (!tl) {
    throw new Error(`Table "${tableModel.name}" is missing top-left cell address (tl)`);
  }
  const startRow = tl.row;
  const startCol = tl.col;
  const endRow = startRow + tableModel.rows.length; // header row + data rows
  const endCol = startCol + columnNames.length - 1;

  const shortRange = colCache.encode(startRow, startCol, endRow, endCol);

  // Use the worksheet name (not table name) for pivotCacheDefinition's worksheetSource
  // The sheet attribute in worksheetSource must reference the actual worksheet name
  const worksheetName = table.worksheet.name;
  const tableName = tableModel.name;

  return {
    name: worksheetName,
    tableName,
    getRow(rowNumber: number): { values: unknown[] } {
      if (rowNumber === 1) {
        return { values: headerRow };
      }
      const dataIndex = rowNumber - DATA_START_INDEX; // rowNumber 2 maps to index 0
      if (dataIndex >= 0 && dataIndex < dataRows.length) {
        return { values: dataRows[dataIndex] };
      }
      return { values: [] };
    },
    getColumn(columnNumber: number): { values: unknown[] } {
      // Validate column number is within bounds
      if (columnNumber < 1 || columnNumber > columnNames.length) {
        return { values: [] };
      }
      // Values should be sparse array with header at index 1, data starting at index 2
      const values: unknown[] = [];
      values[1] = columnNames[columnNumber - 1];
      for (let i = 0; i < tableModel.rows.length; i++) {
        values[i + 2] = tableModel.rows[i][columnNumber - 1];
      }
      return { values };
    },
    getSheetValues(): unknown[][] {
      // Return sparse array where index 1 is header row, and subsequent indices are data rows
      const result: unknown[][] = [];
      result[1] = headerRow;
      for (let i = 0; i < dataRows.length; i++) {
        result[i + 2] = dataRows[i];
      }
      return result;
    },
    dimensions: { shortRange }
  };
}

/** Base cache ID starts at 10 (Excel convention), each subsequent table increments */
const BASE_CACHE_ID = 10;

/**
 * Resolves the data source from the model, supporting both sourceSheet and sourceTable.
 * Validates that exactly one source is provided.
 */
function resolveSource(model: PivotTableModel): PivotTableSource {
  if (model.sourceSheet && model.sourceTable) {
    throw new PivotTableError("Cannot specify both sourceSheet and sourceTable. Choose one.");
  }
  if (model.sourceTable) {
    return createTableSourceAdapter(model.sourceTable);
  }
  if (!model.sourceSheet) {
    throw new Error("Either sourceSheet or sourceTable must be provided.");
  }
  return model.sourceSheet;
}

/** Resolve a value entry to its column name string */
function resolveValueName(v: string | PivotTableValue): string {
  return typeof v === "string" ? v : v.name;
}

/** Resolve a value entry's metric (or undefined if inheriting table-wide default) */
function resolveValueMetric(v: string | PivotTableValue): PivotTableSubtotal | undefined {
  return typeof v === "string" ? undefined : v.metric;
}

function makePivotTable(
  worksheet: { workbook: { pivotTables: PivotTable[] } },
  model: PivotTableModel
): PivotTable {
  // Resolve source (validates exactly one source is provided)
  const source = resolveSource(model);

  validate(model, source);

  const { rows } = model;
  const columns = model.columns ?? [];
  const pages = model.pages ?? [];
  const valueNames = model.values.map(resolveValueName);
  const defaultMetric: PivotTableSubtotal = model.metric ?? "sum";
  const valueMetrics = model.values.map(v => resolveValueMetric(v) ?? defaultMetric);

  const cacheFields = makeCacheFields(source, [...rows, ...columns, ...pages], valueNames);

  const nameToIndex: Record<string, number> = {};
  for (let i = 0; i < cacheFields.length; i++) {
    nameToIndex[cacheFields[i].name] = i;
  }
  const resolveIndex = (fieldName: string, role: string): number => {
    const idx = nameToIndex[fieldName];
    if (idx === undefined) {
      throw new Error(`${role} field "${fieldName}" not found in cache fields`);
    }
    return idx;
  };
  const rowIndices = rows.map(row => resolveIndex(row, "Row"));
  const columnIndices = columns.map(column => resolveIndex(column, "Column"));
  const valueIndices = valueNames.map(value => resolveIndex(value, "Value"));
  const pageIndices = pages.map(page => resolveIndex(page, "Page"));

  // R9-B1: Calculate tableNumber as max(existing)+1 to avoid collision with loaded
  // pivot tables that may have non-contiguous numbering (e.g. [1, 2, 5]).
  // Using length+1 would collide when adding multiple new tables.
  const existingTableNumbers = worksheet.workbook.pivotTables
    .map(pt => pt.tableNumber)
    .filter(n => Number.isFinite(n));
  const tableNumber =
    existingTableNumbers.length > 0
      ? existingTableNumbers.reduce((a, b) => (a > b ? a : b), -Infinity) + 1
      : 1;

  // Dynamic cacheId: avoid collision with existing (loaded) pivot tables.
  // Find the max cacheId already in use and start from max+1.
  const existingCacheIds = worksheet.workbook.pivotTables
    .map(pt => parseInt(pt.cacheId, 10))
    .filter(id => Number.isFinite(id));
  // R8-O3: Use reduce instead of Math.max(...spread) to avoid stack overflow with many pivot tables
  const nextCacheId =
    existingCacheIds.length > 0
      ? existingCacheIds.reduce((a, b) => (a > b ? a : b), -Infinity) + 1
      : BASE_CACHE_ID + tableNumber - 1;

  // form pivot table object
  return {
    source,
    rows: rowIndices,
    columns: columnIndices,
    values: valueIndices,
    pages: pageIndices,
    metric: defaultMetric,
    valueMetrics,
    cacheFields,
    cacheId: String(nextCacheId),
    // Control whether pivot table style overrides worksheet column widths
    // '0' = preserve worksheet column widths (useful for custom sizing)
    // '1' = apply pivot table style width/height (default Excel behavior)
    applyWidthHeightFormats: model.applyWidthHeightFormats ?? "1",
    // Table number for file naming (pivotTable1.xml, pivotTable2.xml, etc.)
    tableNumber
  };
}

function validate(model: PivotTableModel, source: PivotTableSource): void {
  const columns = model.columns ?? [];
  const pages = model.pages ?? [];
  const valueNames = model.values.map(resolveValueName);

  // Get header names from source (already resolved)
  const headerNames = source.getRow(1).values.slice(1);

  // Validate no empty header names (null, undefined, empty-string, or whitespace-only).
  // Note: numeric 0 and boolean false are valid headers (they coerce to "0"/"false").
  for (let i = 0; i < headerNames.length; i++) {
    const h = headerNames[i];
    if (h === null || h === undefined || h === "" || (typeof h === "string" && h.trim() === "")) {
      throw new Error(
        `Empty or missing header name at column ${i + 1} in ${source.name}. Pivot tables require all columns to have non-empty headers.`
      );
    }
  }

  // Validate no duplicate header names
  const headerDupCheck = new Set<string>();
  for (const h of headerNames) {
    const name = String(h);
    if (headerDupCheck.has(name)) {
      throw new Error(
        `Duplicate header name "${name}" found in ${source.name}. Pivot tables require unique column names.`
      );
    }
    headerDupCheck.add(name);
  }

  // Use Set for O(1) lookup — coerce to String for consistent comparison with user-supplied names
  const headerNameSet = new Set(headerNames.map(String));
  const validateFieldExists = (name: string): void => {
    if (!headerNameSet.has(name)) {
      throw new PivotTableError(`The header name "${name}" was not found in ${source.name}.`);
    }
  };
  for (const name of model.rows) {
    validateFieldExists(name);
  }
  for (const name of columns) {
    validateFieldExists(name);
  }
  for (const name of valueNames) {
    validateFieldExists(name);
  }
  for (const name of pages) {
    validateFieldExists(name);
  }

  // Validate no duplicate field names across axis areas (rows, columns, pages).
  // A field can only belong to one axis area. Values can overlap with axis areas (dataField="1").
  const fieldToAxis = new Map<string, string>();
  const axisAreas = [
    { name: "rows", fields: model.rows },
    { name: "columns", fields: columns },
    { name: "pages", fields: pages }
  ];
  for (const area of axisAreas) {
    for (const field of area.fields) {
      const existing = fieldToAxis.get(field);
      if (existing === area.name) {
        throw new Error(
          `Duplicate field "${field}" in ${area.name}. Each field can only appear once per axis area.`
        );
      } else if (existing) {
        throw new Error(
          `Field "${field}" cannot appear in both ${existing} and ${area.name}. Each field can only be assigned to one axis area.`
        );
      }
      fieldToAxis.set(field, area.name);
    }
  }

  if (!model.rows.length) {
    throw new PivotTableError("No pivot table rows specified.");
  }

  // Allow empty columns - Excel will use "Values" as column field
  if (model.values.length < 1) {
    throw new PivotTableError("Must have at least one value.");
  }

  // Validate metric values at runtime (guards against `as any` bypasses)
  if (model.metric !== undefined && !VALID_SUBTOTALS.has(model.metric)) {
    throw new Error(
      `Invalid metric "${model.metric}". Must be one of: ${[...VALID_SUBTOTALS].join(", ")}.`
    );
  }
  for (const v of model.values) {
    const perMetric = resolveValueMetric(v);
    if (perMetric !== undefined && !VALID_SUBTOTALS.has(perMetric)) {
      throw new Error(
        `Invalid metric "${perMetric}" on value field "${resolveValueName(v)}". Must be one of: ${[...VALID_SUBTOTALS].join(", ")}.`
      );
    }
  }

  // Validate no duplicate value field names
  const valueDupCheck = new Set<string>();
  for (const name of valueNames) {
    if (valueDupCheck.has(name)) {
      throw new Error(`Duplicate value field "${name}". Each value field name must be unique.`);
    }
    valueDupCheck.add(name);
  }

  // Multiple values with columns is supported — the synthetic "Values" pseudo-field
  // (field x="-2") is appended to colFields so Excel positions the data field labels
  // correctly on the column axis.
}

/**
 * R8-B2: Unwrap complex cell value types to a primitive suitable for pivot table shared items.
 * - CellErrorValue ({error:"#REF!"}) → PivotErrorValue
 * - CellRichTextValue ({richText:[...]}) → concatenated plain text
 * - CellFormulaValue / CellArrayFormulaValue / CellSharedFormulaValue → result value (recursive)
 * - CellHyperlinkValue ({text,hyperlink}) → text string
 * - CellCheckboxValue ({checkbox:bool}) → boolean
 * - Other objects → String(v)
 */
function unwrapCellValue(v: unknown): string | number | boolean | Date | PivotErrorValue | null {
  if (v === null || v === undefined) {
    return null;
  }
  if (typeof v !== "object") {
    // Already a primitive (string, number, boolean)
    return v as string | number | boolean;
  }
  if (v instanceof Date) {
    return v;
  }
  if (v instanceof PivotErrorValue) {
    return v;
  }
  const obj = v as Record<string, unknown>;
  // CellErrorValue: { error: "#REF!" } → PivotErrorValue with code "REF!" (strip leading #)
  if (typeof obj.error === "string") {
    const errorStr = obj.error as string;
    return new PivotErrorValue(errorStr.startsWith("#") ? errorStr.slice(1) : errorStr);
  }
  // CellFormulaValue / CellArrayFormulaValue / CellSharedFormulaValue: { formula/sharedFormula, result }
  if ("formula" in obj || "sharedFormula" in obj) {
    const result = obj.result;
    if (result === undefined || result === null) {
      return null;
    }
    // result can be number | string | boolean | Date | CellErrorValue — recurse
    return unwrapCellValue(result);
  }
  // CellRichTextValue: { richText: [{text:"..."}, ...] }
  if (Array.isArray(obj.richText)) {
    return (obj.richText as Array<{ text?: string }>).map(rt => rt.text ?? "").join("");
  }
  // CellHyperlinkValue: { text, hyperlink }
  if (typeof obj.text === "string" && typeof obj.hyperlink === "string") {
    return obj.text;
  }
  // CellCheckboxValue: { checkbox: boolean }
  if (typeof obj.checkbox === "boolean") {
    return obj.checkbox;
  }
  // Unknown object — fallback to string
  return String(v);
}

function makeCacheFields(
  source: PivotTableSource,
  fieldNamesWithSharedItems: string[],
  valueFieldNames: string[]
): CacheField[] {
  // Cache fields are used in pivot tables to reference source data.
  // Fields in fieldNamesWithSharedItems get their unique values extracted as sharedItems.
  // Fields in valueFieldNames (but not in fieldNamesWithSharedItems) get min/max calculated.
  // Other fields are unused and get null sharedItems.

  const names = source.getRow(1).values;
  // Use Set for O(1) lookup instead of object
  const sharedItemsFields = new Set(fieldNamesWithSharedItems);
  const valueFields = new Set(valueFieldNames);

  const aggregate = (columnIndex: number): SharedItemValue[] => {
    const columnValues = source.getColumn(columnIndex).values;
    // Build unique values set directly, skipping header (index 0,1).
    // null/undefined are tracked separately (collapsed to a single null sentinel)
    // because Set treats each null as the same key but undefined as a separate key,
    // and we need null in sharedItems so renderCellNew can find it via indexOf.
    const uniqueValues = new Set<SharedItemValue>();
    // R8-B12: PivotErrorValue uses reference equality in Set, so track error strings
    // separately to avoid duplicates (e.g., two PivotErrorValue("#REF!") instances).
    const seenErrors = new Map<string, PivotErrorValue>();
    let hasNull = false;
    for (let i = DATA_START_INDEX; i < columnValues.length; i++) {
      // R8-B2: Unwrap complex cell values (formula results, rich text, errors, etc.)
      const v = unwrapCellValue(columnValues[i]);
      if (v === null || (typeof v === "number" && isNaN(v))) {
        hasNull = true;
      } else if (v instanceof PivotErrorValue) {
        // R8-B12: Deduplicate PivotErrorValue by error string
        if (!seenErrors.has(v.code)) {
          seenErrors.set(v.code, v);
          uniqueValues.add(v);
        }
      } else {
        uniqueValues.add(v);
      }
    }
    const sorted = toSortedArray(uniqueValues);
    // Append null at the end (OOXML convention: <m/> items go last in sharedItems)
    if (hasNull) {
      sorted.push(null);
    }
    return sorted;
  };

  // Calculate min/max and integer status for numeric fields
  const getMinMax = (
    columnIndex: number
  ): { minValue: number; maxValue: number; allInteger: boolean } | null => {
    const columnValues = source.getColumn(columnIndex).values;
    let min = Infinity;
    let max = -Infinity;
    let hasNumeric = false;
    let allInteger = true;
    for (let i = DATA_START_INDEX; i < columnValues.length; i++) {
      // R8-B10: Unwrap formula/complex cell values to extract numeric results
      const unwrapped = unwrapCellValue(columnValues[i]);
      if (typeof unwrapped === "number" && !isNaN(unwrapped)) {
        hasNumeric = true;
        if (unwrapped < min) {
          min = unwrapped;
        }
        if (unwrapped > max) {
          max = unwrapped;
        }
        if (!Number.isInteger(unwrapped)) {
          allInteger = false;
        }
      }
    }
    return hasNumeric ? { minValue: min, maxValue: max, allInteger } : null;
  };

  // Build result array
  const result: CacheField[] = [];
  for (const columnIndex of range(1, names.length)) {
    const rawName = names[columnIndex];
    const name = String(rawName);
    if (sharedItemsFields.has(name)) {
      // Field used for rows/columns - extract unique values as sharedItems
      result.push({ name, sharedItems: aggregate(columnIndex) });
    } else if (valueFields.has(name)) {
      // Field used only for values (aggregation) - calculate min/max
      const minMax = getMinMax(columnIndex);
      result.push({
        name,
        sharedItems: null,
        containsNumber: minMax ? "1" : undefined,
        minValue: minMax?.minValue,
        maxValue: minMax?.maxValue,
        containsInteger: minMax?.allInteger ? "1" : undefined
      });
    } else {
      // Unused field - just empty sharedItems (like Excel does)
      result.push({ name, sharedItems: null });
    }
  }
  return result;
}

export { makePivotTable };
