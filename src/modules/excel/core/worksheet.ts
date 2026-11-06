import {
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable,
  seriesFromColumns as chartSeriesFromColumns
} from "@excel/chart/build/chart-api";
import type {
  AddChartExFromRowsOptions,
  AddChartExFromTableOptions,
  AddChartFromRowsOptions,
  AddChartFromTableOptions,
  SeriesFromColumnsOptions
} from "@excel/chart/build/chart-api";
import { buildChartModel, buildComboChartModel } from "@excel/chart/build/chart-builder";
import { buildChartExModel } from "@excel/chart/build/chart-ex-builder";
import { registerChart, registerChartEx } from "@excel/chart/chart-handle";
import type { AddChartExOptions, ChartExType } from "@excel/chart/model/chart-ex-types";
import { applyChartExPreset, applyChartPreset } from "@excel/chart/model/chart-presets";
import type { ExcelChartExPreset, ExcelChartPreset } from "@excel/chart/model/chart-presets";
import type {
  AddBarChartOptions,
  AddChartOptions,
  AddChartRange,
  AddComboChartOptions,
  AddPieChartOptions,
  AddScatterChartOptions,
  AddSurfaceChartOptions,
  ChartAnchorModel
} from "@excel/chart/model/types";
import type { CellData, FormulaResult, FormulaValueData } from "@excel/core/cell";
import {
  cellAlignment,
  cellCol,
  cellComment,
  cellReadValue,
  cellIsMerged,
  cellMaster,
  cellMerge,
  cellSetComment,
  cellSetValue,
  cellType,
  cellUnmerge,
  cellView,
  checkCellValue
} from "@excel/core/cell";
import type { ColumnData, ColumnModel, ColumnDefn, ColumnView } from "@excel/core/column";
import { columnDefn, columnToModel } from "@excel/core/column";
import type { DataValidationModel } from "@excel/core/data-validations";
import { createDataValidations } from "@excel/core/data-validations";
import { definedNamesHasName } from "@excel/core/defined-names";
import { Enums } from "@excel/core/enums";
import type {
  FormCheckboxData,
  FormCheckboxModel,
  FormCheckboxOptions,
  FormControlRange,
  LegacyControlModel
} from "@excel/core/form-control";
import { formCheckboxCreate, markCheckboxRemoved } from "@excel/core/form-control";
import type { WorksheetImage, ImageModel, ImageModelInput } from "@excel/core/image";
import { applyImageModel, imageClone, imageCreate, imageModel } from "@excel/core/image";
import type { OpaqueRelationship } from "@excel/core/opaque-part";
import {
  isPivotChartSource,
  isPivotFilterTarget,
  withPivotChartSource
} from "@excel/core/pivot-chart";
import { checkNotPivotSource, checkPivotPlacement } from "@excel/core/pivot-source";
import type { PivotTable, PivotTableModel } from "@excel/core/pivot-table";
import { makePivotTable } from "@excel/core/pivot-table";
import type { RangeData, RangeInput } from "@excel/core/range";
import { rangeCreate, rangeExpand, rangeIntersects, rangeRange } from "@excel/core/range";
import type { RowData, RowModel } from "@excel/core/row";
import {
  rowCellCount,
  rowDimensions,
  rowFindCell,
  rowGetModel,
  rowHidden,
  rowValues
} from "@excel/core/row";
import type { AddSparklineGroupOptions, SparklineGroup } from "@excel/core/sparkline";
import { buildSparklineGroup } from "@excel/core/sparkline";
// Chart runtime is imported directly (static). The chart modules depend only
// on the `*-core` data layer (never on this heavy `worksheet.ts`), so the
// dependency graph stays acyclic: `worksheet → chart → *-core`. A consumer
// that never references a chart API gets the entire chart implementation
// tree-shaken out by the bundler — no host registry / install step required.
import { sheetNamesInTabOrder, withReferenceEdit } from "@excel/core/structural-edit";
import type { StyledBlankRangeModel } from "@excel/core/styled-blanks";
import type { TableData, TableModel } from "@excel/core/table";
import {
  checkTableNameFree,
  markTableRemoved,
  tableGetModel,
  tableName,
  tablePrepare,
  tableStore
} from "@excel/core/table";
import type { Workbook } from "@excel/core/workbook";
import {
  getImage,
  releaseChartMedia,
  releaseMedia,
  removeChartEntry,
  removeChartExStructuredEntry,
  refuseLongSheetName,
  removeWorksheetEx,
  validateSheetName
} from "@excel/core/workbook-core";
import type {
  AutoFilterCriteria,
  WorksheetData,
  SheetProtection,
  ChartHandle
} from "@excel/core/worksheet-core";
import {
  _copyStyle,
  _setStyleOption,
  checkRowNumber,
  checkWorksheet,
  checkWorksheetAttached,
  columnCreate,
  columnSetDefn,
  eachRow,
  findCell,
  findRow,
  getCell,
  getColumn,
  get_lastRowNumber,
  getRow,
  getRows,
  getSheetName,
  getSheetWorkbook,
  getTable,
  rowEachCell,
  rowGetCell,
  rowSetValues,
  rowSplice
} from "@excel/core/worksheet-core";
import {
  ExcelColumnOutOfBoundsError,
  ExcelArgumentError,
  ExcelImageError,
  ExcelMergeConflictError,
  ExcelPivotTableError,
  ExcelTableError
} from "@excel/errors";
import type {
  AddImageRange,
  AddShapeOptions,
  AutoFilter,
  CellValue,
  ColBreak,
  ConditionalFormattingOptions,
  HeaderFooter,
  HeaderFooterImagePosition,
  IgnoredError,
  PageSetup,
  RowBreak,
  RowValues,
  ShapeModel,
  Style,
  TableProperties,
  ThreadedComment,
  WatermarkOptions,
  WorksheetProperties,
  WorksheetState,
  WorksheetView
} from "@excel/types";
import type { Origin } from "@excel/utils/address";
import { decodeCell, decodeRange, encodeCol } from "@excel/utils/address";
import { argumentChecks, checkNumber } from "@excel/utils/arguments";
import { getCellDisplayText } from "@excel/utils/cell-format";
import type { DecodedRange } from "@excel/utils/col-cache";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import { isExternalImage } from "@excel/utils/drawing-utils";
import { applyMergeBorders, collectMergeBorders } from "@excel/utils/merge-borders";
import type { LineEdit } from "@excel/utils/reference-rewrite";
import {
  buildSheetProtection,
  splitPasswordOptions,
  verifySheetPassword
} from "@excel/utils/sheet-protection";
import {
  calculateAutoFitWidth,
  getMaxDigitWidth,
  charWidthToPixel,
  getPixelPadding,
  getCellTextWidthPx,
  getCellHeightPt
} from "@excel/utils/text-metrics";
import { excelNameKey } from "@utils/name-key";
import { isRecordArgument } from "@utils/object";

// Type for data validation model - maps address to validation

interface WorksheetOptions {
  workbook?: Workbook;
  id?: number;
  orderNo?: number;
  name?: string;
  state?: WorksheetState;
  properties?: Partial<WorksheetProperties>;
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  views?: Partial<WorksheetView>[];
  autoFilter?: AutoFilter | null;
}

/** The round-trip worksheet model — `Worksheet.getModel` / `Worksheet.setModel`. */
export interface WorksheetModel {
  /**
   * Where this sheet sits on the tab bar, counted across worksheets *and* chartsheets.
   *
   * Present on the model because both writers need it to order the tabs, and its absence was a real defect
   * rather than a gap: a worksheet compared as "unknown position" while a chartsheet carried a number, so a
   * chartsheet added last was drawn first. See `core/sheet-order`.
   */
  orderNo?: number;
  /**
   * Styled blank cells, as rectangles, from a collapsed XLSB read.
   *
   * A cell that carries formatting and no value is one `BrtCellBlank` record, and Excel writes one per cell of a
   * formatted region — thousands for a formatted column, which is the commonest reason a small XLSB is expensive
   * to read. `blankCells: "collapse"` accumulates them here instead of materialising each, and the XLSB writer
   * expands them again, so the round trip is unchanged and no fidelity report is owed.
   */
  styledBlankRanges?: readonly StyledBlankRangeModel[];
  /**
   * Notes a *streamed* sheet kept, because it no longer has the rows they were attached to.
   *
   * `readonly unknown[]` for the same reason the reader's side of `styledBlankRanges` was: the element is the XLSB
   * writer's `SheetComment`, and naming it here would make this model depend on that writer. See `getSheetModel`.
   */
  retainedComments?: readonly unknown[];
  /** Hyperlinks a *streamed* sheet kept, with the relationship ids its part already names — see `retainedComments`. */
  retainedHyperlinks?: readonly unknown[];
  id: number;
  name: string;
  dataValidations: DataValidationModel;
  properties: Partial<WorksheetProperties>;
  state: WorksheetState;
  pageSetup: PageSetup;
  headerFooter: HeaderFooter;
  rowBreaks: RowBreak[];
  colBreaks: ColBreak[];
  views: Partial<WorksheetView>[];
  autoFilter: AutoFilter | null;
  autoFilterCriteria?: AutoFilterCriteria;
  sortStateXml?: string;
  /**
   * Relationships from this sheet to preserved parts this library does not
   * model. See `WorksheetData._opaqueRels`.
   */
  opaqueRels?: OpaqueRelationship[];
  /** See `WorksheetData._legacyControls`. */
  legacyControls?: LegacyControlModel[];
  /** See `WorksheetData._chartsheetPlaceholder`. */
  chartsheetPlaceholder?: string;
  /** See `WorksheetData._xlsbDrawingRelationshipId`. */
  xlsbDrawingRelationshipId?: string;
  worksheetNamespaceAttributes?: Record<string, string>;
  worksheetMcIgnorable?: string;
  sortStateAutoFilterRef?: string;
  media: ImageModel[];
  shapes?: ShapeModel[];
  sheetProtection: SheetProtection | null;
  tables: TableModel[];
  pivotTables: PivotTable[];
  conditionalFormattings: ConditionalFormattingOptions[];
  formControls: FormCheckboxModel[];
  ignoredErrors: IgnoredError[];
  watermark?: WatermarkOptions | null;
  /** Internal snapshot used to make API-owned header watermark removal reversible. */
  cols?: ColumnModel[];
  rows?: RowModel[];
  dimensions?: RangeData;
  mergeCells?: string[];
  /** Loaded drawing data (for charts, etc.) - preserved for round-trip */
  drawing?: unknown;
  /** Chart anchor models for worksheet charts */
  charts?: ChartAnchorModel[];
  /** Sparkline groups (x14:sparklineGroups) */
  sparklineGroups?: SparklineGroup[];
  /**
   * Office 365 threaded comments for this worksheet. Rendered as a
   * separate `xl/threadedComments/threadedComment{N}.xml` part. Empty
   * when the sheet has no modern comments.
   */
  threadedComments?: Array<{ ref: string; comment: ThreadedComment }>;
}

// Worksheet requirements
//  Operate as sheet inside workbook or standalone
//  Load and Save from file and stream
//  Access/Add/Delete individual cells
//  Manage column widths and row heights

export function createWorksheet(options: WorksheetOptions): WorksheetData {
  const ws = {} as WorksheetData;

  ws._workbook = options.workbook!;

  // in a workbook, each sheet will have a number
  ws.id = options.id ?? 0;
  ws.orderNo = options.orderNo ?? 0;

  // and a name - use the setter to ensure validation and truncation
  assignSheetName(ws, options.name || `sheet${ws.id}`);

  // add a state
  ws.state = options.state ?? "visible";

  // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
  // Note: _rows is zero based. Must subtract 1 to go from Cell.row(cell) to index
  ws._rows = [];

  // column definitions
  ws._columns = [];

  // column keys (addRow convenience): key ==> this._collumns index
  // No prototype, so a key such as `"__proto__"` or `"toString"` is an ordinary entry.
  ws._keys = Object.create(null) as Record<string, ColumnData>;

  // keep record of all merges
  ws._merges = {};

  // record of all row and column pageBreaks
  ws.rowBreaks = [];
  ws.colBreaks = [];

  // for tabColor, default row height, outline levels, etc
  ws.properties = {
    defaultRowHeight: 15,
    outlineLevelCol: 0,
    outlineLevelRow: 0,
    ...options.properties
  };

  // for all things printing
  ws.pageSetup = {
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    orientation: "portrait",
    // Excel does not normally write these unless explicitly set.
    // Historically we used 4294967295 as a sentinel when parsing, but emitting it
    // can cause strict Excel parsers to treat the workbook as corrupted.
    horizontalDpi: undefined,
    verticalDpi: undefined,
    fitToPage: !!(
      options.pageSetup &&
      (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
      !options.pageSetup.scale
    ),
    pageOrder: "downThenOver",
    blackAndWhite: false,
    draft: false,
    cellComments: "None",
    errors: "displayed",
    scale: 100,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: undefined,
    showRowColHeaders: false,
    showGridLines: false,
    firstPageNumber: undefined,
    horizontalCentered: false,
    verticalCentered: false,
    rowBreaks: null,
    colBreaks: null,
    ...options.pageSetup
  } as PageSetup;

  ws.headerFooter = {
    differentFirst: false,
    differentOddEven: false,
    scaleWithDoc: true,
    alignWithMargins: true,
    oddHeader: null,
    oddFooter: null,
    evenHeader: null,
    evenFooter: null,
    firstHeader: null,
    firstFooter: null,
    ...options.headerFooter
  };

  ws.dataValidations = createDataValidations(undefined, ws);

  // for freezepanes, split, zoom, gridlines, etc
  ws.views = options.views ?? [];

  ws.autoFilter = options.autoFilter ?? null;

  // for images, etc
  ws._media = [];

  // for user-drawn shapes (rectangles, lines, text boxes, …)
  ws._shapes = [];

  // for charts
  ws._charts = [];
  ws._sparklineGroups = [];

  // worksheet protection
  ws.sheetProtection = null;

  // for tables
  ws.tables = {};

  ws.pivotTables = [];

  ws.conditionalFormattings = [];

  // for form controls (legacy checkboxes, etc.)
  ws.formControls = [];

  // ignored errors (suppress green triangles in Excel)
  ws.ignoredErrors = [];

  // Office 365 threaded comments (separate from classic VML notes).
  ws.threadedComments = [];

  // watermark configuration
  ws._watermark = null;
  ws._watermarkMedia = null;

  return ws;
}

/**
 * Remove the sheet from its workbook. References to it from what remains become `#REF!`, and a 3-D span that ends on
 * it shrinks to the next sheet inward, as Excel does.
 */
export function destroy(ws: WorksheetData): void {
  checkWorksheet(ws);
  const wb = ws._workbook;
  if (!wb || wb._worksheets[ws.id] !== ws) {
    removeWorksheetEx(wb, ws);
    return;
  }
  checkNotPivotSource(wb, { sheet: ws }, "Workbook.removeWorksheet");
  const edit = { kind: "removeSheet", sheet: ws._name, sheets: sheetNamesInTabOrder(wb) } as const;
  withReferenceEdit(wb, edit, () => removeWorksheetEx(wb, ws));
}

export function spliceColumns(
  ws: WorksheetData,
  start: number,
  count: number,
  ...inserts: CellValue[][]
): void {
  checkLineSplice(ws, "Column.splice", "columns", start, count);
  inserts.forEach((column, i) => {
    if (!Array.isArray(column)) {
      throw new ExcelArgumentError(
        `inserts[${i}]`,
        `Column.splice: inserts[${i}] must be an array of the column's values, got ${typeof column}.`
      );
    }
    // Every value before any cell moves: a refusal found while the rows were being spliced left the rows above it
    // shifted, the rows below it not, and every reference already rewritten.
    column.forEach((value, r) =>
      checkCellValue(value, { where: "Column.splice", field: `inserts[${i}][${r}]` })
    );
  });
  withLineEdit(ws, "columns", start, count, inserts.length, () =>
    _spliceColumnCells(ws, start, count, inserts)
  );
}

/**
 * The checks a row or column splice makes before anything moves. A negative count used to run the insertion branch
 * with a deletion's arithmetic — duplicating a line and shifting every reference by one — and `NaN` or a fraction
 * reached the cell arrays as an index nothing can hold. A removed sheet is refused like every other edit: the
 * reference rewrite is keyed by sheet name, so a splice through a stale handle rewrote the references of whichever
 * sheet had taken the name since.
 */
function checkLineSplice(
  ws: WorksheetData,
  where: string,
  kind: "rows" | "columns",
  start: number,
  count: number
): void {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  // A start outside the grid is a position error, as for every other row or column reference.
  if (kind === "rows") {
    checkRowNumber(start);
  } else if (!Number.isInteger(start) || start < 1 || start > 16384) {
    throw new ExcelColumnOutOfBoundsError(start);
  }
  checkNumber(where, "count", count, { min: 0, integer: true });
}

/**
 * Run a row or column splice as a structural edit, so every reference to the sheet's cells follows them.
 *
 * A splice overwrites the lines it both deletes and inserts — they stay where they are, with new values — and only
 * the difference is an insertion or a deletion: `splice(3, 2, a)` writes `a` over line 3 and deletes line 4. That is
 * the edit references see, so a formula naming line 3 keeps naming it, as it would after `Cell.setValue`.
 */
function withLineEdit(
  ws: WorksheetData,
  kind: "rows" | "columns",
  start: number,
  count: number,
  inserted: number,
  apply: () => void
): void {
  const overwritten = Math.min(count, inserted);
  const wb = getSheetWorkbook(ws);
  if (count === inserted || !wb) {
    apply();
    return;
  }
  const edit: LineEdit = {
    kind,
    sheet: getSheetName(ws),
    at: start + overwritten,
    deleted: count - overwritten,
    inserted: inserted - overwritten
  };
  withReferenceEdit(wb, edit, apply, { overwritten });
}

function _spliceColumnCells(
  ws: WorksheetData,
  start: number,
  count: number,
  inserts: CellValue[][]
): void {
  // Before splicing cells, release all cell-level merge references so that
  // row.splice copies plain values instead of merge proxies.
  // _spliceMerges (called later) will rebuild cell-level refs at new coordinates.
  for (const merge of Object.values(ws._merges)) {
    for (let r = merge.top; r <= merge.bottom; r++) {
      for (let c = merge.left; c <= merge.right; c++) {
        const cell = findCell(ws, r, c);
        if (cell && cellType(cell) === Enums.ValueType.Merge) {
          cellUnmerge(cell);
        }
      }
    }
  }

  if (inserts.length > 0) {
    // Every row the sheet has or an inserted column reaches, whether it exists yet or not: stopping at the sheet's last
    // row dropped the values below it — all of them on an empty sheet.
    const nRows = Math.max(ws._rows.length, ...inserts.map(insert => insert.length));
    for (let i = 0; i < nRows; i++) {
      const insertValues = inserts.map(insert => insert[i] ?? null);
      const row = getRow(ws, i + 1);
      rowSplice(row, start, count, ...insertValues);
    }
  } else {
    // nothing to insert, so just splice all rows
    ws._rows.forEach(r => {
      if (r) {
        rowSplice(r, start, count);
      }
    });
  }

  // splice column definitions
  const nExpand = inserts.length - count;
  const nKeep = start + count;
  const nEnd = ws._columns.length;
  if (nExpand < 0) {
    for (let i = start + inserts.length; i <= nEnd; i++) {
      columnSetDefn(getColumn(ws, i), columnDefn(getColumn(ws, i - nExpand)));
    }
  } else if (nExpand > 0) {
    for (let i = nEnd; i >= nKeep; i--) {
      columnSetDefn(getColumn(ws, i + nExpand), columnDefn(getColumn(ws, i)));
    }
  }
  for (let i = start; i < start + inserts.length; i++) {
    columnSetDefn(getColumn(ws, i), undefined);
  }

  // account for merges
  _spliceMerges(ws, "col", start, count, inserts.length);
}

export function insertRow(
  ws: WorksheetData,
  pos: number,
  value: RowValues,
  style: string = "n"
): RowData {
  checkWorksheet(ws);
  spliceRows(ws, pos, 0, value);
  _setStyleOption(ws, pos, style);
  return getRow(ws, pos);
}

export function insertRows(
  ws: WorksheetData,
  pos: number,
  values: RowValues[],
  style: string = "n"
): RowData[] | undefined {
  checkWorksheet(ws);
  spliceRows(ws, pos, 0, ...values);
  if (style !== "n") {
    // copy over the styles
    for (let i = 0; i < values.length; i++) {
      if (style[0] === "o" && findRow(ws, values.length + pos + i) !== undefined) {
        _copyStyle(ws, values.length + pos + i, pos + i, style[1] === "+");
      } else if (style[0] === "i" && findRow(ws, pos - 1) !== undefined) {
        _copyStyle(ws, pos - 1, pos + i, style[1] === "+");
      }
    }
  }
  return getRows(ws, pos, values.length);
}

export function duplicateRow(
  ws: WorksheetData,
  rowNum: number,
  count: number,
  insert: boolean = false
): void {
  // create count duplicates of rowNum
  // either inserting new or overwriting existing rows
  checkNumber("Row.duplicate", "count", count, { min: 0, integer: true });

  const rSrc = getRow(ws, rowNum);
  const inserts = Array.from<RowValues>({ length: count }).fill(rowValuesForStructuralCopy(rSrc));

  // Collect single-row merges from the source row before splicing
  // (only merges where top == bottom == rowNum, i.e. horizontal merges within one row)
  const srcMerges: RangeData[] = [];
  for (const merge of Object.values(ws._merges)) {
    if (merge.top === rowNum && merge.bottom === rowNum) {
      srcMerges.push(merge);
    }
  }

  // Collect images anchored to the source row before splicing
  // (images whose top-left anchor is on the source row)
  const srcImages: WorksheetImage[] = [];
  const srcRow0 = rowNum - 1; // 0-based source row
  for (const image of ws._media) {
    if (image.type === "image" && image.range) {
      if (image.range.tl.nativeRow === srcRow0) {
        srcImages.push(image);
      }
    }
  }

  spliceRows(ws, rowNum + 1, insert ? 0 : count, ...inserts);

  // now copy styles...
  for (let i = 0; i < count; i++) {
    const rDst = ws._rows[rowNum + i];
    rDst.style = copyStyle(rSrc.style) ?? {};
    rDst.height = rSrc.height;
    rowEachCell(rSrc, { includeEmpty: true }, (cell: CellData, colNumber: number) => {
      rowGetCell(rDst, colNumber).style = copyStyle(cell.style) ?? {};
    });
  }

  // Duplicate single-row merges from source row into each new row
  if (srcMerges.length > 0) {
    for (let i = 0; i < count; i++) {
      const dstRow = rowNum + 1 + i;

      // In overwrite mode, clear any existing merges in the target row
      if (!insert) {
        const toRemove: string[] = [];
        for (const [key, merge] of Object.entries(ws._merges)) {
          if (merge.top <= dstRow && merge.bottom >= dstRow) {
            toRemove.push(key);
          }
        }
        for (const key of toRemove) {
          _unMergeMaster(ws, getCell(ws, key));
        }
      }

      for (const srcMerge of srcMerges) {
        mergeCellsWithoutStyle(ws, dstRow, srcMerge.left, dstRow, srcMerge.right);
      }
    }
  }

  // Duplicate images from source row into each new row.
  // In overwrite mode, first remove any images anchored to the target rows
  // so they don't coexist with the clones (mirrors merge cleanup above).
  if (!insert) {
    const dstStart0 = rowNum; // first target row, 0-based (1-based rowNum + 1 → 0-based rowNum)
    const dstEnd0 = rowNum + count - 1; // last target row, 0-based
    ws._media = ws._media.filter(image => {
      if (image.type === "image" && image.range) {
        const row0 = image.range.tl.nativeRow;
        if (row0 >= dstStart0 && row0 <= dstEnd0) {
          releaseMedia(ws._workbook, image.imageId);
          return false;
        }
      }
      return true;
    });
  }

  for (let i = 0; i < count; i++) {
    const rowDelta = i + 1; // offset from source row to target row
    for (const srcImage of srcImages) {
      const cloned = imageClone(srcImage);
      cloned.range!.tl.nativeRow = srcRow0 + rowDelta;
      if (cloned.range!.br) {
        const brDelta = srcImage.range!.br!.nativeRow - srcRow0;
        cloned.range!.br.nativeRow = srcRow0 + rowDelta + brDelta;
      }
      ws._media.push(cloned);
    }
  }
}

export function spliceRows(
  ws: WorksheetData,
  start: number,
  count: number,
  ...inserts: RowValues[]
): void {
  checkLineSplice(ws, "Row.splice", "rows", start, count);
  withLineEdit(ws, "rows", start, count, inserts.length, () =>
    _spliceRowCells(ws, start, count, inserts)
  );
}

function _spliceRowCells(
  ws: WorksheetData,
  start: number,
  count: number,
  inserts: RowValues[]
): void {
  // same problem as row.splice, except worse.

  // Before splicing rows, release all cell-level merge references so that
  // row value copies work on plain values instead of merge proxies.
  // _spliceMerges (called later) will rebuild cell-level refs at new coordinates.
  for (const merge of Object.values(ws._merges)) {
    for (let r = merge.top; r <= merge.bottom; r++) {
      for (let c = merge.left; c <= merge.right; c++) {
        const cell = findCell(ws, r, c);
        if (cell && cellType(cell) === Enums.ValueType.Merge) {
          cellUnmerge(cell);
        }
      }
    }
  }

  const nKeep = start + count;
  const nInserts = inserts.length;
  const nExpand = nInserts - count;
  const nEnd = ws._rows.length;
  let i: number;
  let rSrc: RowData | undefined;
  if (nExpand < 0) {
    // remove rows
    // A block that reaches past the last row has nothing below it to move up, so the loop below never runs:
    // clear what it deletes here. The rows it overwrites with `inserts` are written further down.
    if (nKeep > nEnd) {
      for (i = start + nInserts; i <= nEnd; i++) {
        ws._rows[i - 1] = undefined!;
      }
    }
    for (i = nKeep; i <= nEnd; i++) {
      rSrc = ws._rows[i - 1];
      if (rSrc) {
        const rDst = getRow(ws, i + nExpand);
        rowSetValues(rDst, rowValuesForStructuralCopy(rSrc));
        rDst.style = (copyStyle(rSrc.style) as Partial<Style>) ?? {};
        rDst.height = rSrc.height;
        rowEachCell(rSrc, { includeEmpty: true }, (cell: CellData, colNumber: number) => {
          const target = rowGetCell(rDst, colNumber);
          target.style = (copyStyle(cell.style) as Partial<Style>) ?? {};
          cellSetComment(target, cellComment(cell));
        });
        ws._rows[i - 1] = undefined!;
      } else {
        ws._rows[i + nExpand - 1] = undefined!;
      }
    }
  } else if (nExpand > 0) {
    // insert new cells
    for (i = nEnd; i >= nKeep; i--) {
      rSrc = ws._rows[i - 1];
      if (rSrc) {
        const rDst = getRow(ws, i + nExpand);
        rowSetValues(rDst, rowValuesForStructuralCopy(rSrc));
        rDst.style = (copyStyle(rSrc.style) as Partial<Style>) ?? {};
        rDst.height = rSrc.height;
        rowEachCell(rSrc, { includeEmpty: true }, (cell: CellData, colNumber: number) => {
          const target = rowGetCell(rDst, colNumber);
          target.style = (copyStyle(cell.style) as Partial<Style>) ?? {};
          cellSetComment(target, cellComment(cell));
        });
      } else {
        ws._rows[i + nExpand - 1] = undefined!;
      }
    }
  }

  // now copy over the new values
  for (i = 0; i < nInserts; i++) {
    const rDst = getRow(ws, start + i);
    rDst.style = {};
    rowSetValues(rDst, inserts[i]);
  }

  // account for merges
  _spliceMerges(ws, "row", start, count, nInserts);
}

function rowValuesForStructuralCopy(row: RowData): CellValue[] {
  const values = rowValues(row);
  row.cells.forEach(cell => {
    if (cell?._formulaGhostOwner !== undefined) {
      values[cellCol(cell)] = null;
    }
  });
  return values;
}

export function mergeCells(ws: WorksheetData, ...cells: RangeInput[]): void {
  const dimensions = rangeCreate(cells);
  _mergeCellsInternal(ws, dimensions);
}

export function mergeCellsWithoutStyle(ws: WorksheetData, ...cells: RangeInput[]): void {
  const dimensions = rangeCreate(cells);
  _mergeCellsInternal(ws, dimensions, true);
}

function _mergeCellsInternal(
  ws: WorksheetData,
  dimensions: RangeData,
  ignoreStyle?: boolean
): void {
  // check cells aren't already merged
  Object.values(ws._merges).forEach((merge: RangeData) => {
    if (rangeIntersects(merge, dimensions)) {
      throw new ExcelMergeConflictError();
    }
  });

  const { top, left, bottom, right } = dimensions;

  // Collect perimeter borders BEFORE merge overwrites slave styles
  const collected = ignoreStyle
    ? undefined
    : collectMergeBorders(top, left, bottom, right, (r, c) => findCell(ws, r, c));

  // Apply merge — slave cells inherit the master's full style
  const master = getCell(ws, dimensions.top, dimensions.left);
  // A merge turns the top-left cell into user-owned content. It must no
  // longer be treated as a derived spill ghost during cleanup/serialization.
  master._formulaGhostOwner = undefined;
  for (let i = top; i <= bottom; i++) {
    for (let j = left; j <= right; j++) {
      if (i > top || j > left) {
        cellMerge(getCell(ws, i, j), master, ignoreStyle);
      }
    }
  }

  // Reconstruct position-aware borders (like Excel):
  // outer borders survive, inner borders are cleared.
  if (collected) {
    applyMergeBorders(top, left, bottom, right, collected, (r, c) => getCell(ws, r, c));
  }

  // index merge
  ws._merges[master.address] = dimensions;
}

function _unMergeMaster(ws: WorksheetData, master: CellData): void {
  // master is always top left of a rectangle
  const merge = ws._merges[master.address];
  if (merge) {
    for (let i = merge.top; i <= merge.bottom; i++) {
      for (let j = merge.left; j <= merge.right; j++) {
        cellUnmerge(getCell(ws, i, j));
      }
    }
    delete ws._merges[master.address];
  }
}

function _spliceMerges(
  ws: WorksheetData,
  axis: "row" | "col",
  start: number,
  count: number,
  nInserts: number
): void {
  const nExpand = nInserts - count;
  if (nExpand === 0 && count === 0) {
    return;
  }
  const nKeep = start + count;
  const isRow = axis === "row";

  const newMerges: Record<string, RangeData> = {};

  for (const merge of Object.values(ws._merges)) {
    const { top, left, bottom, right } = merge;
    // For row axis: lo=top, hi=bottom. For col axis: lo=left, hi=right.
    const lo = isRow ? top : left;
    const hi = isRow ? bottom : right;

    if (nExpand <= 0 && count > 0) {
      // Deleting rows/columns
      const deleteEnd = nKeep - 1;
      if (lo > deleteEnd) {
        // Entirely after deleted range — shift
        const newRange = isRow
          ? rangeCreate(top + nExpand, left, bottom + nExpand, right)
          : rangeCreate(top, left + nExpand, bottom, right + nExpand);
        newMerges[colCache.encodeAddress(newRange.top, newRange.left)] = newRange;
      } else if (hi < start) {
        // Entirely before deleted range — unchanged
        newMerges[colCache.encodeAddress(top, left)] = merge;
      } else if (lo >= start && hi <= deleteEnd) {
        // Entirely within deleted range — remove
      } else {
        // Spans splice boundary — shrink
        let newTop = top;
        let newLeft = left;
        let newBottom = bottom;
        let newRight = right;
        if (isRow) {
          newTop = top < start ? top : start;
          newBottom = Math.max(newTop, bottom + nExpand);
        } else {
          newLeft = left < start ? left : start;
          newRight = Math.max(newLeft, right + nExpand);
        }
        const newRange = rangeCreate(newTop, newLeft, newBottom, newRight);
        if (newTop === newBottom && newLeft === newRight) {
          // Degenerate 1x1 merge — remove instead of keeping
        } else {
          newMerges[colCache.encodeAddress(newRange.top, newRange.left)] = newRange;
        }
      }
    } else {
      // Inserting rows/columns: shift items at/after nKeep
      if (lo >= nKeep) {
        // Entirely at or after splice — shift
        const newRange = isRow
          ? rangeCreate(top + nExpand, left, bottom + nExpand, right)
          : rangeCreate(top, left + nExpand, bottom, right + nExpand);
        newMerges[colCache.encodeAddress(newRange.top, newRange.left)] = newRange;
      } else if (hi < nKeep) {
        // Entirely before splice — unchanged
        newMerges[colCache.encodeAddress(top, left)] = merge;
      } else {
        // Spans splice boundary — stretch
        if (isRow) {
          merge.bottom = bottom + nExpand;
        } else {
          merge.right = right + nExpand;
        }
        newMerges[colCache.encodeAddress(top, left)] = merge;
      }
    }
  }

  ws._merges = newMerges;

  // Rebuild cell-level merge references for all merges.
  // Pre-unmerge in spliceRows/spliceColumns clears all cell refs,
  // so we must rebuild every merge, not just moved/resized ones.
  for (const m of Object.values(newMerges)) {
    const master = getCell(ws, m.top, m.left);
    for (let r = m.top; r <= m.bottom; r++) {
      for (let c = m.left; c <= m.right; c++) {
        if (r > m.top || c > m.left) {
          cellMerge(getCell(ws, r, c), master, true);
        }
      }
    }
  }
}

export function unMergeCells(ws: WorksheetData, ...cells: RangeInput[]): void {
  checkWorksheetAttached(ws);
  const dimensions = rangeCreate(cells);

  // find any cells in that range and unmerge them
  for (let i = dimensions.top; i <= dimensions.bottom; i++) {
    for (let j = dimensions.left; j <= dimensions.right; j++) {
      const cell = findCell(ws, i, j);
      if (cell) {
        if (cellType(cell) === Enums.ValueType.Merge) {
          // this cell merges to another master
          _unMergeMaster(ws, cellMaster(cell));
        } else if (ws._merges[cell.address]) {
          // this cell is a master
          _unMergeMaster(ws, cell);
        }
      }
    }
  }
}

/** What {@link fillFormula} takes beside the range and the formula. */
export interface FillFormulaOptions {
  /**
   * The cached result of each cell, as a grid (`results[row][col]`), a flat row-major list, or a function of the
   * cell's row and column numbers. Without it the cells carry no cached value until the workbook is calculated.
   */
  readonly results?:
    | FormulaResult[][]
    | FormulaResult[]
    | ((row: number, col: number) => FormulaResult | undefined);
  /**
   * `"shared"` (the default) writes the formula once and shares it to every other cell of the range, each cell's
   * references shifting with it. `"array"` writes one array formula spanning the range.
   */
  readonly shareType?: "shared" | "array";
}

/** Fill `range` with `formula`, written once on its top-left cell and shared to the rest (or as one array formula). */
export function fillFormula(
  ws: WorksheetData,
  range: string,
  formula: string,
  options?: FillFormulaOptions
): void;
/**
 * @deprecated Pass the results and the share type in the options object:
 * `fillFormula(ws, range, formula, { results, shareType })`. Removed in 1.0.
 */
export function fillFormula(
  ws: WorksheetData,
  range: string,
  formula: string,
  results?:
    | FormulaResult[][]
    | FormulaResult[]
    | ((row: number, col: number) => FormulaResult | undefined),
  shareType?: string
): void;
export function fillFormula(
  ws: WorksheetData,
  range: string,
  formula: string,
  resultsOrOptions?:
    | FillFormulaOptions
    | FormulaResult[][]
    | FormulaResult[]
    | ((row: number, col: number) => FormulaResult | undefined),
  positionalShareType?: string
): void {
  let results: FillFormulaOptions["results"];
  let shareType: string;
  if (typeof resultsOrOptions !== "function" && isRecordArgument(resultsOrOptions)) {
    results = resultsOrOptions.results;
    shareType = resultsOrOptions.shareType ?? "shared";
  } else {
    results = resultsOrOptions;
    shareType = positionalShareType ?? "shared";
  }
  // Define formula for top-left cell and share to rest
  const decoded = colCache.decode(range) as DecodedRange;
  const { top, left, bottom, right } = decoded;
  const width = right - left + 1;
  const masterAddress = colCache.encodeAddress(top, left);
  const isShared = shareType === "shared";

  // work out result accessor
  let getResult: (row: number, col: number) => FormulaResult | undefined;
  if (typeof results === "function") {
    getResult = results;
  } else if (Array.isArray(results)) {
    if (Array.isArray(results[0])) {
      getResult = (row: number, col: number) =>
        (results as FormulaResult[][])[row - top][col - left];
    } else {
      getResult = (row: number, col: number) =>
        (results as FormulaResult[])[(row - top) * width + (col - left)];
    }
  } else {
    getResult = () => undefined;
  }
  let first = true;
  for (let r = top; r <= bottom; r++) {
    for (let c = left; c <= right; c++) {
      if (first) {
        const cell = getCell(ws, r, c);
        const formulaValue: FormulaValueData = {
          shareType,
          formula,
          ref: range,
          result: getResult(r, c)
        };
        cellSetValue(cell, formulaValue as CellValue);
        first = false;
      } else {
        cellSetValue(
          getCell(ws, r, c),
          isShared
            ? {
                sharedFormula: masterAddress,
                result: getResult(r, c)
              }
            : getResult(r, c)
        );
      }
    }
  }
}

export function addImage(
  ws: WorksheetData,
  imageId: string | number,
  range: AddImageRange
): WorksheetImage {
  checkWorksheetAttached(ws);
  const model = {
    type: "image",
    imageId: String(imageId),
    range
  };
  const image = imageCreate(ws, model);
  ws._media.push(image);
  return image;
}

export function addShape(ws: WorksheetData, options: AddShapeOptions): ShapeModel {
  checkWorksheetAttached(ws);
  const range = options.range;
  // A shape must cover an area, mirroring images. Reject inputs that resolve
  // to no size up front, with a clear shape-specific message — otherwise the
  // failure surfaces much later as a confusing `ExcelImageError` from the internal
  // range parser when the worksheet is serialized.
  const hasArea =
    (typeof range === "string" && range.includes(":")) ||
    (typeof range === "object" &&
      range !== null &&
      ("br" in range || "ext" in range || "pos" in range));
  if (!hasArea) {
    throw new ExcelImageError(
      'addShape requires a range covering an area: a cell range like "B2:D5", or an object with `br`, `ext`, or `pos`.'
    );
  }
  const shape: ShapeModel = {
    type: "shape",
    shapeType: options.type ?? "rect",
    range,
    fillColor: options.fillColor,
    lineColor: options.lineColor,
    lineWidth: options.lineWidth,
    text: options.text,
    name: options.name
  };
  ws._shapes.push(shape);
  return shape;
}

export function getShapes(ws: WorksheetData): ShapeModel[] {
  return ws._shapes.slice();
}

/** Remove a shape — one {@link getShapes} listed, or its index in that list. Returns `true` if one was removed. */
export function removeShape(ws: WorksheetData, shape: ShapeModel | number): boolean {
  checkWorksheetAttached(ws);
  const index = listPosition(ws._shapes, shape);
  if (index < 0) {
    return false;
  }
  ws._shapes.splice(index, 1);
  return true;
}

function _resolveShapeModel(ws: WorksheetData, shape: ShapeModel): ShapeModel {
  let range: Extract<ImageModel, { type: "image" }>["range"] | undefined;
  try {
    const probe = imageCreate(ws, { type: "image", imageId: "", range: shape.range });
    // The probe is always an "image" type, so its model carries `range`.
    range = (imageModel(probe) as Extract<ImageModel, { type: "image" }>).range;
  } catch {
    // Range could not be parsed into an anchor (addShape validates the common
    // cases up front; this guards exotic inputs). Drop the anchor so the
    // serializer skips this shape rather than failing the whole worksheet.
    range = undefined;
  }
  if (!range) {
    return { ...shape, anchorRange: undefined };
  }
  return {
    ...shape,
    anchorRange: {
      tl: range.tl,
      br: range.br,
      ext: range.ext,
      pos: range.pos,
      editAs: range.editAs
    }
  };
}

export function getImages(ws: WorksheetData): WorksheetImage[] {
  return ws._media.filter(m => m.type === "image");
}

/**
 * The position `item` names in `list`: the index itself when it is an integer inside the list, the
 * item's own index otherwise, `-1` when neither. A fractional or out-of-range position names nothing —
 * `splice` would have truncated `1.5` to `1` and removed an item nobody asked for.
 */
function listPosition<T>(list: readonly T[], item: T | number): number {
  if (typeof item === "number") {
    return Number.isInteger(item) && item >= 0 && item < list.length ? item : -1;
  }
  return list.indexOf(item);
}

/**
 * Replace a placed image's model — {@link applyImageModel}, recording that the picture it showed is no
 * longer used here when the model names a different one.
 */
export function setImageModel(image: WorksheetImage, model: ImageModelInput): void {
  checkWorksheetAttached(image.worksheet);
  const previous = image.imageId;
  applyImageModel(image, model);
  if (previous !== undefined && previous !== image.imageId) {
    releaseMedia(image.worksheet._workbook, previous);
  }
}

/**
 * Remove a placed image from the sheet — the handle `getImages` returned, or its
 * index in that list. Returns `true` if an image was removed.
 *
 * Only the placement goes: the picture stays registered in the workbook, so its
 * image id can still be placed elsewhere. A watermark is not a placed image and
 * is removed with `removeWatermark`.
 */
export function removeImage(ws: WorksheetData, image: WorksheetImage | number): boolean {
  checkWorksheetAttached(ws);
  const images = getImages(ws);
  const position = listPosition(images, image);
  const target = position < 0 ? undefined : images[position];
  if (!target || target.type !== "image") {
    return false;
  }
  const idx = ws._media.indexOf(target);
  if (idx < 0) {
    return false;
  }
  ws._media.splice(idx, 1);
  releaseMedia(ws._workbook, target.imageId);
  return true;
}

export function addChart(
  ws: WorksheetData,
  options: AddChartOptions,
  range: AddChartRange
): number {
  return registerChart(ws, buildChartModel(options), range, options);
}

export function addColumnChart(
  ws: WorksheetData,
  options: Omit<AddBarChartOptions, "type" | "barDir">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "bar", barDir: "col" }, range);
}

export function addBarChart(
  ws: WorksheetData,
  options: Omit<AddBarChartOptions, "type" | "barDir">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "bar", barDir: "bar" }, range);
}

export function addLineChart(
  ws: WorksheetData,
  options: Omit<AddChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "line" }, range);
}

export function addAreaChart(
  ws: WorksheetData,
  options: Omit<AddChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "area" }, range);
}

export function addPieChart(
  ws: WorksheetData,
  options: Omit<AddPieChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "pie" }, range);
}

export function addDoughnutChart(
  ws: WorksheetData,
  options: Omit<AddPieChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "doughnut" }, range);
}

export function addScatterChart(
  ws: WorksheetData,
  options: Omit<AddScatterChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "scatter" }, range);
}

export function addBubbleChart(
  ws: WorksheetData,
  options: Omit<AddChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "bubble" }, range);
}

export function addRadarChart(
  ws: WorksheetData,
  options: Omit<AddChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "radar" }, range);
}

export function addStockChart(
  ws: WorksheetData,
  options: Omit<AddChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "stock" }, range);
}

export function addSurfaceChart(
  ws: WorksheetData,
  options: Omit<AddSurfaceChartOptions, "type">,
  range: AddChartRange
): number {
  return addChart(ws, { ...options, type: "surface" }, range);
}

export function addHistogramChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "histogram" }, range);
}

export function addParetoChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "pareto" }, range);
}

export function addWaterfallChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "waterfall" }, range);
}

export function addFunnelChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "funnel" }, range);
}

export function addTreemapChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "treemap" }, range);
}

export function addSunburstChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "sunburst" }, range);
}

export function addBoxWhiskerChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "boxWhisker" }, range);
}

export function addRegionMapChart(
  ws: WorksheetData,
  options: Omit<AddChartExOptions, "type">,
  range: AddChartRange
): number {
  return addChartEx(ws, { ...options, type: "regionMap" }, range);
}

export function addPresetChart(
  ws: WorksheetData,
  preset: ExcelChartPreset,
  options: Omit<AddChartOptions, "type"> & Partial<Pick<AddChartOptions, "type">>,
  range: AddChartRange
): number {
  return addChart(ws, applyChartPreset(preset, options), range);
}

export function addPresetChartEx(
  ws: WorksheetData,
  preset: ExcelChartExPreset,
  options: Omit<AddChartExOptions, "type"> & Partial<Pick<AddChartExOptions, "type">>,
  range: AddChartRange
): number {
  return addChartEx(ws, applyChartExPreset(preset, options), range);
}

/**
 * A chart series over ranges of `ws`: `values` (and `categories`, when given) qualified with the sheet's name, so
 * `"B2:B9"` becomes `"'My Sheet'!$B$2:$B$9"`. A range that already names a sheet keeps that sheet, made absolute the
 * same way.
 */
export function seriesFromColumns(
  ws: WorksheetData,
  options: SeriesFromColumnsOptions
): ReturnType<typeof chartSeriesFromColumns> {
  return chartSeriesFromColumns(getSheetName(ws), options);
}

export function addChartFromTable(
  ws: WorksheetData,
  table: TableData | string,
  options: AddChartFromTableOptions,
  range: AddChartRange
): number {
  return addChart(ws, chartOptionsFromTable(ws, table, options), range);
}

export function addChartFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartFromRowsOptions<T>,
  range: AddChartRange
): number {
  return addChart(ws, chartOptionsFromRows(ws, rows, options), range);
}

export function addColumnChartFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: Omit<AddChartFromRowsOptions<T>, "type" | "barDir">,
  range: AddChartRange
): number {
  return addChart(
    ws,
    chartOptionsFromRows(ws, rows, {
      ...options,
      type: "bar",
      barDir: "col"
    } as AddChartFromRowsOptions<T>),
    range
  );
}

export function addChartExFromTable(
  ws: WorksheetData,
  table: TableData | string,
  options: AddChartExFromTableOptions & {
    type: Exclude<AddChartExOptions["type"], "regionMap">;
  },
  range: AddChartRange
): number {
  return addChartEx(ws, chartExOptionsFromTable(ws, table, options), range);
}

export function addChartExFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartExFromRowsOptions<T> & {
    type: Exclude<AddChartExOptions["type"], "regionMap">;
  },
  range: AddChartRange
): number {
  return addChartEx(ws, chartExOptionsFromRows(ws, rows, options), range);
}

export function addPivotChart(
  ws: WorksheetData,
  pivotTable: PivotTable,
  options: AddChartOptions,
  range: AddChartRange
): number {
  const pivotChartOptions = withPivotChartSource(pivotTable, options);
  return registerChart(ws, buildChartModel(pivotChartOptions), range, pivotChartOptions);
}

export function addPivotComboChart(
  ws: WorksheetData,
  pivotTable: PivotTable,
  options: AddComboChartOptions,
  range: AddChartRange
): number {
  const pivotChartOptions = withPivotChartSource(pivotTable, options);
  return registerChart(ws, buildComboChartModel(pivotChartOptions), range, pivotChartOptions);
}

export function addComboChart(
  ws: WorksheetData,
  options: AddComboChartOptions,
  range: AddChartRange
): number {
  return registerChart(ws, buildComboChartModel(options), range, options);
}

export function addChartEx(
  ws: WorksheetData,
  options: AddChartExOptions,
  range: AddChartRange
): number {
  return registerChartEx(ws, buildChartExModel(options), range);
}

// A record rather than a list so that a `ChartExType` added later fails to compile here until it is
// listed — a list only rejects names that do not exist, not ones that are missing.
const CHART_EX_TYPES: ReadonlySet<string> = new Set(
  Object.keys({
    sunburst: true,
    treemap: true,
    waterfall: true,
    funnel: true,
    histogram: true,
    pareto: true,
    boxWhisker: true,
    regionMap: true
  } satisfies Record<ChartExType, true>)
);

/**
 * What {@link placeChart} draws. The shape of the options decides the chart family:
 * `groups` makes a combo chart, a `type` from the Office 2016 family (`histogram`, `pareto`,
 * `waterfall`, `funnel`, `treemap`, `sunburst`, `boxWhisker`, `regionMap`) a chartEx part, any
 * other `type` a classic chart. Adding `pivotTable` binds a classic or combo chart to that pivot.
 */
export type ChartPlacement =
  | AddChartOptions
  | AddComboChartOptions
  | AddChartExOptions
  | ((AddChartOptions | AddComboChartOptions) & { pivotTable: PivotTable });

/**
 * Add a chart to the sheet and return its handle — the value every other chart operation,
 * `removeChart` included, takes.
 *
 * Every chart the `add*Chart` family can make, this makes: those functions are this one with
 * the options' `type` filled in, and the preset / table / rows variants build their options with
 * the public option builders (`applyPreset`, `optionsFromTable`, `optionsFromRows`) first.
 *
 * @example
 * ```typescript
 * const chart = Chart.place(ws, "E1:L14", { type: "bar", barDir: "col", series });
 * Chart.setStyle(chart, 26);
 * Chart.remove(ws, chart);
 * ```
 */
/** Whether `type` names one of the Office 2016 chart types, which are written as ChartEx parts. */
export function isChartExType(type: unknown): boolean {
  return typeof type === "string" && CHART_EX_TYPES.has(type);
}

export function placeChart(
  ws: WorksheetData,
  options: ChartPlacement,
  range: AddChartRange
): ChartHandle {
  checkWorksheetAttached(ws);
  // `pivotTable: undefined` — an optional field spread through from elsewhere — binds nothing.
  if ("pivotTable" in options && options.pivotTable !== undefined) {
    const { pivotTable, ...rest } = options;
    if ("groups" in rest) {
      addPivotComboChart(ws, pivotTable, rest, range);
    } else {
      addPivotChart(ws, pivotTable, rest, range);
    }
  } else if ("groups" in options) {
    addComboChart(ws, options as AddComboChartOptions, range);
  } else if (CHART_EX_TYPES.has(options.type)) {
    addChartEx(ws, options as AddChartExOptions, range);
  } else {
    addChart(ws, options as AddChartOptions, range);
  }
  // Every chart-adding path appends exactly one handle.
  return ws._charts[ws._charts.length - 1];
}

/**
 * The sheet's charts, in the order they were added. A chart-adding call always appends, so
 * `getCharts(ws).at(-1)` is the handle of the chart just added.
 */
export function getCharts(ws: WorksheetData): ChartHandle[] {
  return [...ws._charts];
}

/**
 * Remove a chart from the sheet. Returns `true` if one was removed.
 *
 * `chart` is a handle — from `placeChart` or `getCharts` — or the chart's **position in
 * `getCharts(ws)`**. It is never the number an `add*Chart` call returned: that is a workbook-wide
 * part number (`chartN.xml` / `chartExN.xml`) and matches the position only by accident.
 */
export function removeChart(ws: WorksheetData, chart: ChartHandle | number): boolean {
  checkWorksheetAttached(ws);
  const idx = listPosition(ws._charts, chart);
  if (idx < 0) {
    return false;
  }
  const removed = ws._charts.splice(idx, 1)[0];
  if (removed.chartNumber > 0) {
    releaseChartMedia(ws._workbook, removed.chartNumber);
    removeChartEntry(ws._workbook, removed.chartNumber);
  }
  if (removed.chartExNumber > 0) {
    removeChartExStructuredEntry(ws._workbook, removed.chartExNumber);
  }
  // Prune the matching anchor from the loaded drawing so the writer
  // doesn't emit a dangling rel pointing at the now-removed chart
  // part. The drawing xform regenerates rels from `drawing.anchors`
  // on every write (see `worksheet-xform.ts` chart-anchor reconcile),
  // so dropping the anchor here is sufficient — we don't need to
  // hand-edit `drawing.rels` ourselves.
  const drawing = ws._drawing as
    | { anchors?: Array<{ chartNumber?: number; chartExNumber?: number }> }
    | undefined;
  if (drawing?.anchors) {
    drawing.anchors = drawing.anchors.filter(anchor => {
      if (removed.chartNumber > 0 && anchor.chartNumber === removed.chartNumber) {
        return false;
      }
      if (removed.chartExNumber > 0 && anchor.chartExNumber === removed.chartExNumber) {
        return false;
      }
      return true;
    });
  }
  return true;
}

export function addSparklineGroup(
  ws: WorksheetData,
  options: AddSparklineGroupOptions
): SparklineGroup {
  checkWorksheetAttached(ws);
  const group = buildSparklineGroup(options);
  ws._sparklineGroups.push(group);
  return group;
}

/**
 * The sheet's sparkline groups, in the order they were added — a new array of the groups the sheet holds and writes,
 * so editing one changes the sheet without the checks `Sparkline.add` makes.
 */
export function getSparklineGroups(ws: WorksheetData): SparklineGroup[] {
  return [...ws._sparklineGroups];
}

/** Remove a sparkline group — one `Sparkline.add` returned or `getSparklineGroups` listed, or its index there. */
export function removeSparklineGroup(
  ws: WorksheetData,
  groupOrIndex: SparklineGroup | number
): boolean {
  checkWorksheetAttached(ws);
  const idx = listPosition(ws._sparklineGroups, groupOrIndex);
  if (idx < 0) {
    return false;
  }
  ws._sparklineGroups.splice(idx, 1);
  return true;
}

/**
 * Set the sheet's background picture, replacing the one it had. A sheet has one background: a second
 * call used to add another entry, which `getBackgroundImageId` never reported and the writer wrote
 * over the first.
 */
export function addBackgroundImage(ws: WorksheetData, imageId: string | number): void {
  checkWorksheetAttached(ws);
  const bookImage = getImage(ws._workbook, imageId);
  if (bookImage && isExternalImage(bookImage)) {
    throw new ExcelImageError(
      "Background images cannot be external (linked) images. " +
        "Use an embedded image (buffer/base64/filename). " +
        "External images are only supported for cell pictures and overlay watermarks."
    );
  }
  removeBackgroundImage(ws);
  const model = {
    type: "background",
    imageId: String(imageId)
  };
  ws._media.push(imageCreate(ws, model));
}

/** Remove the sheet's background picture. Returns `true` if it had one. */
export function removeBackgroundImage(ws: WorksheetData): boolean {
  checkWorksheetAttached(ws);
  const backgrounds = ws._media.filter(medium => medium.type === "background");
  if (backgrounds.length === 0) {
    return false;
  }
  ws._media = ws._media.filter(medium => medium.type !== "background");
  backgrounds.forEach(medium => releaseMedia(ws._workbook, medium.imageId));
  return true;
}

export function getBackgroundImageId(ws: WorksheetData): string | undefined {
  const image = ws._media.find(m => m.type === "background");
  return image && image.imageId;
}

export function addWatermark(ws: WorksheetData, options: WatermarkOptions): void {
  checkWorksheetAttached(ws);
  const mode = options.mode ?? "overlay";

  // Validate BEFORE mutating any state: VML header/footer images use
  // embedded media (`<v:imagedata o:relid>`); external (linked) images are
  // not representable here. Reject them up front so a failed call leaves the
  // existing watermark untouched (no partial mutation).
  if (mode === "header") {
    const bookImage = getImage(ws._workbook, options.imageId);
    if (bookImage && isExternalImage(bookImage)) {
      throw new ExcelImageError(
        "Header watermark images cannot be external (linked) images. " +
          "Use an embedded image (buffer/base64/filename), or use overlay mode for linked images."
      );
    }
  }

  // Replace only the watermark owned by this API. Imported header/footer
  // images are independent document content and must survive both replacement
  // and removal of the API watermark.
  if (ws._watermark?.mode === "header") {
    removeWatermark(ws);
  } else {
    removeOwnedWatermarkMedia(ws);
  }

  const position = options.position ?? "CH";

  ws._watermark = {
    imageId: String(options.imageId),
    mode,
    opacity: options.opacity,
    headerWidth: options.headerWidth,
    headerHeight: options.headerHeight,
    applyTo: options.applyTo,
    position: options.position
  };

  if (ws._watermark.mode === "overlay") {
    // Add as a special "watermark" media entry for the drawing pipeline
    const model = {
      type: "watermark",
      imageId: String(options.imageId),
      opacity: options.opacity
    };
    const medium = imageCreate(ws, model);
    ws._media.push(medium);
    ws._watermarkMedia = medium;
  } else {
    // Delegate section ownership, `&G` bookkeeping, and replacement semantics
    // to the explicit six-section API.
    setHeaderFooterImage(ws, {
      imageId: options.imageId,
      position,
      width: options.headerWidth,
      height: options.headerHeight,
      applyTo: options.applyTo
    });
    ws._watermarkMedia =
      ws._media.find(
        medium => medium.type === "headerImage" && (medium.position ?? "CH") === position
      ) ?? null;
  }
}

/** Section token (`&L`/`&C`/`&R`) that anchors an image position. */
function headerFooterSectionToken(position: HeaderFooterImagePosition): string {
  return position[0] === "L" ? "&L" : position[0] === "R" ? "&R" : "&C";
}

/**
 * Header/footer fields a given image position writes into, ordered
 * odd → even → first, filtered by the caller's `applyTo` selection.
 */
function headerFooterImageFields(
  position: HeaderFooterImagePosition,
  applyTo: "all" | "odd" | "even" | "first",
  headerFooter: HeaderFooter
): Array<keyof HeaderFooter> {
  const isFooter = position[1] === "F";
  const odd = isFooter ? "oddFooter" : "oddHeader";
  const even = isFooter ? "evenFooter" : "evenHeader";
  const first = isFooter ? "firstFooter" : "firstHeader";
  if (applyTo === "odd") {
    return [odd];
  }
  if (applyTo === "even") {
    return [even];
  }
  if (applyTo === "first") {
    return [first];
  }
  return [odd, even, first];
}

/**
 * Remove the `&G` image placeholder from one header/footer section across all
 * three page variants, leaving every other section and code untouched.
 */
function stripSectionImagePlaceholder(
  source: HeaderFooter,
  position: HeaderFooterImagePosition
): HeaderFooter {
  const result: HeaderFooter = { ...source };
  const token = headerFooterSectionToken(position);
  const fields: Array<keyof HeaderFooter> =
    position[1] === "F"
      ? ["oddFooter", "evenFooter", "firstFooter"]
      : ["oddHeader", "evenHeader", "firstHeader"];

  for (const field of fields) {
    const value = result[field];
    if (typeof value !== "string" || !sectionContainsImage(value, token)) {
      continue;
    }
    const start = value.indexOf(token);
    const imageIndex = value.indexOf("&G", start + token.length);
    let cleaned = value.slice(0, imageIndex) + value.slice(imageIndex + 2);
    const sectionStart = cleaned.indexOf(token);
    if (sectionStart >= 0) {
      const sectionTail = cleaned.slice(sectionStart + token.length);
      const nextSection = sectionTail.search(/&[LCR]/);
      const sectionContent = nextSection < 0 ? sectionTail : sectionTail.slice(0, nextSection);
      if (sectionContent === "") {
        cleaned = cleaned.slice(0, sectionStart) + cleaned.slice(sectionStart + token.length);
      }
    }
    result[field] = (cleaned || null) as never;
  }
  return result;
}

function sectionContainsImage(value: string | null | undefined, token: string): boolean {
  if (!value) {
    return false;
  }
  const start = value.indexOf(token);
  if (start < 0) {
    return false;
  }
  const tail = value.slice(start + token.length);
  const nextSection = tail.search(/&[LCR]/);
  return (nextSection < 0 ? tail : tail.slice(0, nextSection)).includes("&G");
}

function insertImagePlaceholder(value: string | null, token: string): string {
  const existing = value ?? "";
  if (sectionContainsImage(existing, token)) {
    return existing;
  }
  const index = existing.indexOf(token);
  return index >= 0
    ? `${existing.slice(0, index + token.length)}&G${existing.slice(index + token.length)}`
    : `${existing}${token}&G`;
}

/** Pure transformation used both when adding and when precisely undoing. */
function buildWatermarkedHeaderFooter(
  source: HeaderFooter,
  position: HeaderFooterImagePosition,
  applyTo: "all" | "odd" | "even" | "first"
): HeaderFooter {
  const result: HeaderFooter = { ...source };
  const token = headerFooterSectionToken(position);
  const isFooter = position[1] === "F";
  const oddField = isFooter ? "oddFooter" : "oddHeader";
  if (applyTo === "even" && !result.differentOddEven) {
    result[isFooter ? "evenFooter" : "evenHeader"] = result[oddField];
    result.differentOddEven = true;
  }
  if (applyTo === "all") {
    if (!result.differentOddEven) {
      result[isFooter ? "evenFooter" : "evenHeader"] = result[oddField];
      result.differentOddEven = true;
    }
    if (!result.differentFirst) {
      result[isFooter ? "firstFooter" : "firstHeader"] = result[oddField];
      result.differentFirst = true;
    }
  }
  if (applyTo === "first" && !result.differentFirst) {
    result[isFooter ? "firstFooter" : "firstHeader"] = result[oddField];
    result.differentFirst = true;
  }
  for (const field of headerFooterImageFields(position, applyTo, result)) {
    const current = result[field];
    if (typeof current === "string" || current === null) {
      result[field] = insertImagePlaceholder(current, token) as never;
    }
  }
  return result;
}

/**
 * Remove the API-owned media entry. Imported sibling images are untouched.
 */
function removeOwnedWatermarkMedia(ws: WorksheetData): void {
  const owned = ws._watermarkMedia;
  if (!owned) {
    return;
  }
  // Identity match: the API removes exactly the entry it created, never an
  // identically-positioned image that came from the source document.
  ws._media = ws._media.filter(medium => medium !== owned);
  releaseMedia(ws._workbook, owned.imageId);
  ws._watermarkMedia = null;
}

/**
 * Options for {@link setHeaderFooterImage}.
 */
export interface HeaderFooterImageOptions {
  /** Image ID obtained from `Workbook.addImage()`. */
  imageId: string | number;
  /** Section that holds the image. */
  position: HeaderFooterImagePosition;
  /** Image width in points. Defaults to Excel's 467.25. */
  width?: number;
  /** Image height in points. Defaults to Excel's 311.25. */
  height?: number;
  /**
   * Which page variants show the image.
   *
   * @default "odd" — the shared field used by every page while the odd/even
   * and first-page variants are disabled.
   */
  applyTo?: "all" | "odd" | "even" | "first";
}

/** A header/footer image as stored on the worksheet. */
export interface HeaderFooterImageEntry {
  imageId: string;
  position: HeaderFooterImagePosition;
  width?: number;
  height?: number;
}

/**
 * Place an image in one of Excel's six header/footer sections
 * (`LH`/`CH`/`RH`/`LF`/`CF`/`RF`), replacing whatever occupied that section.
 *
 * This is the explicit counterpart to the single-value {@link addWatermark}
 * API: it addresses each section independently, so images loaded from an
 * existing workbook can be inspected, replaced, or removed individually.
 */
export function setHeaderFooterImage(ws: WorksheetData, options: HeaderFooterImageOptions): void {
  checkWorksheetAttached(ws);
  const bookImage = getImage(ws._workbook, options.imageId);
  if (bookImage && isExternalImage(bookImage)) {
    throw new ExcelImageError(
      "Header/footer images cannot be external (linked) images. " +
        "Use an embedded image (buffer/base64/filename)."
    );
  }

  removeHeaderFooterImage(ws, options.position);
  ws._media.push(
    imageCreate(ws, {
      type: "headerImage",
      imageId: String(options.imageId),
      headerWidth: options.width,
      headerHeight: options.height,
      applyTo: options.applyTo,
      position: options.position
    })
  );
  ws.headerFooter = buildWatermarkedHeaderFooter(
    ws.headerFooter,
    options.position,
    options.applyTo ?? "odd"
  );
}

/** Every header/footer image on the worksheet, in section order. */
export function getHeaderFooterImages(ws: WorksheetData): HeaderFooterImageEntry[] {
  return ws._media
    .filter(medium => medium.type === "headerImage")
    .map(medium => ({
      // Parsed workbooks carry numeric media indexes; normalise so the public
      // shape is always the string id used by `Workbook.addImage()`.
      imageId: String(medium.imageId ?? ""),
      position: medium.position ?? "CH",
      width: medium.headerWidth,
      height: medium.headerHeight
    }));
}

/**
 * Remove the image in one header/footer section together with its `&G`
 * placeholder. Returns `true` when a section image was removed.
 */
export function removeHeaderFooterImage(
  ws: WorksheetData,
  position: HeaderFooterImagePosition
): boolean {
  checkWorksheetAttached(ws);
  const occupants = ws._media.filter(
    medium => medium.type === "headerImage" && (medium.position ?? "CH") === position
  );
  if (occupants.length === 0) {
    return false;
  }
  ws._media = ws._media.filter(medium => !occupants.includes(medium));
  occupants.forEach(medium => releaseMedia(ws._workbook, medium.imageId));
  ws.headerFooter = stripSectionImagePlaceholder(ws.headerFooter, position);
  if (occupants.includes(ws._watermarkMedia as never)) {
    // The single-value watermark API owned this entry; drop its bookkeeping so
    // a later removeWatermark() cannot resurrect a stale snapshot.
    ws._watermark = null;
    ws._watermarkMedia = null;
  }
  return true;
}

/** The sheet's watermark settings, or `null` when it has none. */
export function getWatermark(ws: WorksheetData): WatermarkOptions | null {
  return ws._watermark;
}

/** Remove the sheet's watermark. Returns `true` if it had one. */
export function removeWatermark(ws: WorksheetData): boolean {
  checkWorksheetAttached(ws);
  const watermark = ws._watermark;
  if (!watermark) {
    return false;
  }
  if (watermark.mode === "header") {
    removeHeaderFooterImage(ws, watermark.position ?? "CH");
  } else {
    removeOwnedWatermarkMedia(ws);
  }
  ws._watermark = null;
  ws._watermarkMedia = null;
  return true;
}

export function addFormCheckbox(
  ws: WorksheetData,
  range: FormControlRange,
  options?: FormCheckboxOptions
): FormCheckboxData {
  checkWorksheetAttached(ws);
  const checkbox = formCheckboxCreate(ws, range, options);
  ws.formControls.push(checkbox);
  return checkbox;
}

/**
 * The sheet's live checkbox array — mutating it changes the sheet. Kept for the deprecated
 * `Form.listCheckboxes`, which callers used to remove a checkbox by splicing it before
 * `removeFormCheckbox` existed; use {@link listFormCheckboxes}.
 */
export function getFormCheckboxes(ws: WorksheetData): FormCheckboxData[] {
  return ws.formControls;
}

/** The sheet's checkboxes, in the order they were added — a copy, like every other `list`. */
export function listFormCheckboxes(ws: WorksheetData): FormCheckboxData[] {
  return [...ws.formControls];
}

/**
 * Remove a form checkbox — the handle `addFormCheckbox` / `getFormCheckboxes`
 * returned, or its index in that list. Returns `true` if one was removed.
 *
 * The remaining checkboxes keep their shape ids, and a checkbox added afterwards
 * is given one above every id still in use, so a removal never makes two
 * controls share an id.
 */
export function removeFormCheckbox(
  ws: WorksheetData,
  checkbox: FormCheckboxData | number
): boolean {
  checkWorksheetAttached(ws);
  const idx = listPosition(ws.formControls, checkbox);
  if (idx < 0) {
    return false;
  }
  const [removed] = ws.formControls.splice(idx, 1);
  markCheckboxRemoved(removed);
  return true;
}

/** Options for {@link protect}: the protection flags, plus the password that unlocks them. */
export interface ProtectSheetOptions extends Partial<SheetProtection> {
  /** Password hashed into the sheet; omit to protect without one. */
  password?: string;
}

/** Protect the sheet, optionally with a password: `protect(ws, { password: "pw", formatCells: true })`. */
export function protect(ws: WorksheetData, options?: ProtectSheetOptions): Promise<void>;
/**
 * @deprecated Pass the password in the options object: `protect(ws, { password, ...options })`.
 * Removed in 1.0.
 */
export function protect(
  ws: WorksheetData,
  password: string | undefined,
  options?: Partial<SheetProtection>
): Promise<void>;
export async function protect(
  ws: WorksheetData,
  passwordOrOptions?: string | ProtectSheetOptions,
  legacyOptions?: Partial<SheetProtection>
): Promise<void> {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  const [password, options] = splitPasswordOptions(passwordOrOptions, legacyOptions);
  ws.sheetProtection = await buildSheetProtection(password, options);
}

/** Remove the sheet's protection, as `protect` set it. */
export function unprotect(ws: WorksheetData): void {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  ws.sheetProtection = null;
}

/**
 * Verify a candidate password against the sheet's stored protection hash
 * (set via `protect()`). Returns `false` when the sheet isn't protected, or
 * was protected without a password to check against.
 */
export async function verifyPassword(ws: WorksheetData, password: string): Promise<boolean> {
  checkWorksheet(ws);
  // Checked as a password is: a number used to reach the hash and fail there with a `TypeError`.
  argumentChecks("Worksheet.verifyPassword").required("password", password);
  argumentChecks("Worksheet.verifyPassword").string("password", password);
  return verifySheetPassword(ws.sheetProtection, password);
}

/**
 * Add a table to the sheet and return its handle.
 *
 * Nothing on the sheet changes until every check has passed: the model is validated, its values are checked as a
 * cell's are, its area must not overlap another table, and its name and display name must be free in the workbook.
 * (`Table.add` refuses a name Excel cannot hold before this runs; the deprecated `Table.create` corrects one.)
 *
 * @throws {ExcelTableError} When a check fails — the name is already taken in the workbook (case-insensitive), or the
 * table would overlap another one.
 */
export function addTable(ws: WorksheetData, model: TableProperties): TableData {
  checkWorksheetAttached(ws);
  // A copy, so the caller's object is not the table's model: editing it afterwards would change the table without
  // any of the table's checks, and validation would otherwise write defaults into it.
  const table = tablePrepare(ws, structuredClone(model) as TableModel);
  // The name as stored: a caller's is valid already (`Table.add` refuses one that is not), a deprecated
  // `Table.create` model's is corrected by `tableValidate`, and the duplicate checks compare that.
  const nameKey = excelNameKey(tableName(table));
  if (ws.tables[tableName(table)]) {
    throw new ExcelTableError(
      `Table name "${tableName(table)}" already exists in worksheet "${getSheetName(ws)}".`
    );
  }
  if (getSheetWorkbook(ws)._tableNames.has(nameKey)) {
    throw new ExcelTableError(
      `Table name "${tableName(table)}" already exists in the workbook. ` +
        `Table names must be unique across the entire workbook, compared as Excel compares them (case-insensitive).`
    );
  }
  if (definedNamesHasName(getSheetWorkbook(ws)._definedNames, tableName(table))) {
    throw new ExcelTableError(
      `Table name "${tableName(table)}" is a defined name in the workbook. ` +
        `Tables and defined names share one namespace, compared as Excel compares them (case-insensitive).`
    );
  }
  checkTableNameFree(table, tableName(table), "name");
  if (table.table.displayName !== undefined) {
    checkTableNameFree(table, table.table.displayName, "displayName");
  }
  tableStore(table);
  ws.tables[tableName(table)] = table;
  getSheetWorkbook(ws)._tableNames.add(nameKey);
  return table;
}

/**
 * Remove a table — by its handle, or by `name`: its stored name or the one it was added under, as
 * {@link getTable} resolves it. Returns `true` if the sheet had it. Its handle refuses edits afterwards.
 *
 * @throws {ExcelPivotTableError} when a pivot table reads its data from the table: the pivot cache would name a table
 * that no longer exists.
 */
export function removeTable(ws: WorksheetData, name: string | TableData): boolean {
  checkWorksheetAttached(ws);
  const table =
    typeof name === "string"
      ? getTable(ws, name)
      : Object.values(ws.tables).find(candidate => candidate === name);
  if (table === undefined) {
    return false;
  }
  checkNotPivotSource(getSheetWorkbook(ws), { table, sheet: ws }, "Table.remove");
  const key = Object.keys(ws.tables).find(candidate => ws.tables[candidate] === table)!;
  getSheetWorkbook(ws)._tableNames.delete(excelNameKey(key));
  delete ws.tables[key];
  markTableRemoved(table);
  return true;
}

export function addPivotTable(ws: WorksheetData, model: PivotTableModel): PivotTable {
  checkWorksheetAttached(ws);
  const pivotTable = makePivotTable(
    { workbook: getSheetWorkbook(ws), name: getSheetName(ws) },
    model
  );
  checkPivotPlacement(ws.pivotTables, pivotTable);

  ws.pivotTables.push(pivotTable);
  getSheetWorkbook(ws).pivotTables.push(pivotTable);

  return pivotTable;
}

/** The pivot tables on the worksheet, in the order they were added. */
export function getPivotTables(ws: WorksheetData): PivotTable[] {
  return [...ws.pivotTables];
}

/**
 * Remove a pivot table from the sheet — the handle `addPivotTable` /
 * `getPivotTables` returned, or its index in that list. Returns `true` if one
 * was removed.
 *
 * Its pivot cache goes with it unless another pivot table shares that cache:
 * caches are written from the pivot tables that remain.
 *
 * A pivot table something else is bound to cannot be removed while that binding
 * exists, since it would point at nothing — this throws a {@link ExcelPivotTableError}:
 * - a pivot chart (remove the chart first, `Chart.remove`);
 * - a slicer or timeline of a loaded workbook. Those are preserved as written and
 *   cannot be edited here, so such a pivot table stays.
 */
export function removePivotTable(ws: WorksheetData, pivot: PivotTable | number): boolean {
  checkWorksheetAttached(ws);
  const idx = listPosition(ws.pivotTables, pivot);
  if (idx < 0) {
    return false;
  }
  const target = ws.pivotTables[idx];
  const wb = getSheetWorkbook(ws);
  if (isPivotChartSource(wb, target)) {
    throw new ExcelPivotTableError(
      `Pivot table "${target.name ?? `PivotTable${target.tableNumber}`}" is the source of a pivot chart; remove the chart first.`
    );
  }
  if (isPivotFilterTarget(wb, target, ws.id)) {
    throw new ExcelPivotTableError(
      `Pivot table "${target.name ?? `PivotTable${target.tableNumber}`}" is filtered by a slicer or timeline; it cannot be removed.`
    );
  }
  ws.pivotTables.splice(idx, 1);
  const bookIdx = wb.pivotTables.indexOf(target);
  if (bookIdx >= 0) {
    wb.pivotTables.splice(bookIdx, 1);
  }
  return true;
}

export function addConditionalFormatting(
  ws: WorksheetData,
  cf: ConditionalFormattingOptions
): ConditionalFormattingOptions {
  checkWorksheetAttached(ws);
  ws.conditionalFormattings.push(cf);
  return cf;
}

/** The sheet's conditional-formatting entries, in the order they were added. */
export function getConditionalFormattings(ws: WorksheetData): ConditionalFormattingOptions[] {
  return [...ws.conditionalFormattings];
}

/** Which conditional-formatting entries {@link removeConditionalFormattingEntry} removes. */
export type ConditionalFormattingTarget =
  | ConditionalFormattingOptions
  | number
  | ((
      value: ConditionalFormattingOptions,
      index: number,
      array: ConditionalFormattingOptions[]
    ) => boolean);

/**
 * Remove one conditional-formatting entry: the entry `addConditionalFormatting` returned (or
 * `getConditionalFormattings` lists), the entry at a position in that list, or the first
 * entry a predicate accepts. Returns `true` if one was removed. To remove every entry, call
 * {@link clearConditionalFormattings}.
 */
export function removeConditionalFormattingEntry(
  ws: WorksheetData,
  target: ConditionalFormattingTarget
): boolean {
  checkWorksheetAttached(ws);
  const entries = ws.conditionalFormattings;
  // A predicate removes the first entry it accepts, as `remove` removes one; `clear` removes them all.
  const index =
    typeof target === "function" ? entries.findIndex(target) : listPosition(entries, target);
  if (index < 0) {
    return false;
  }
  entries.splice(index, 1);
  return true;
}

/** Remove every conditional-formatting entry of the sheet. */
export function clearConditionalFormattings(ws: WorksheetData): void {
  checkWorksheetAttached(ws);
  ws.conditionalFormattings = [];
}

/**
 * Remove conditional formatting: the entry at an index, every entry a predicate
 * matches, or — with no filter — all of them.
 *
 * The form behind the deprecated `Worksheet.removeConditionalFormatting`; the namespace
 * splits it into {@link removeConditionalFormattingEntry} and {@link clearConditionalFormattings}.
 */
export function removeConditionalFormatting(
  ws: WorksheetData,
  filter?:
    | number
    | ((
        value: ConditionalFormattingOptions,
        index: number,
        array: ConditionalFormattingOptions[]
      ) => boolean)
): void {
  if (typeof filter === "number") {
    ws.conditionalFormattings.splice(filter, 1);
  } else if (typeof filter === "function") {
    // Keep entries for which the predicate returns false; drop the matches.
    ws.conditionalFormattings = ws.conditionalFormattings.filter(
      (value, index, array) => !filter(value, index, array)
    );
  } else {
    ws.conditionalFormattings = [];
  }
}

/**
 * Options for {@link autoFitColumn} and {@link autoFitColumns}.
 */
export interface AutoFitColumnOptions {
  /**
   * Also measure cells in hidden rows — for example rows inside a collapsed
   * outline group, so the column is already wide enough once it is expanded.
   *
   * @default false — hidden rows are ignored, as Excel's own AutoFit does.
   */
  includeHiddenRows?: boolean;
}

/**
 * Set a column's width to fit its content, as Excel's AutoFit does.
 *
 * Cells in hidden rows are ignored unless `includeHiddenRows` is set. A
 * merged cell counts as visible while any row of the merge is, since that is
 * where Excel draws it; a merge spanning several columns is not measured,
 * because its width cannot be attributed to one of them.
 */
export function autoFitColumn(
  ws: WorksheetData,
  col: number | string,
  options?: AutoFitColumnOptions
): WorksheetData {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  const colNum = typeof col === "string" ? colCache.l2n(col) : col;
  _autoFitColumnImpl(ws, colNum, options);
  return ws;
}

/**
 * {@link autoFitColumn} for every column in use, or for `startCol`–`endCol`.
 * Pass the options on their own to fit every column:
 * `autoFitColumns(ws, { includeHiddenRows: true })`.
 */
export function autoFitColumns(ws: WorksheetData, options?: AutoFitColumnOptions): WorksheetData;
export function autoFitColumns(
  ws: WorksheetData,
  startCol?: number | string,
  endCol?: number | string,
  options?: AutoFitColumnOptions
): WorksheetData;
export function autoFitColumns(
  ws: WorksheetData,
  startColOrOptions?: number | string | AutoFitColumnOptions,
  endCol?: number | string,
  rangeOptions?: AutoFitColumnOptions
): WorksheetData {
  checkWorksheetAttached(ws);
  const rangeGiven = startColOrOptions === null || typeof startColOrOptions !== "object";
  const startCol = rangeGiven ? startColOrOptions : undefined;
  const options = rangeGiven ? rangeOptions : startColOrOptions;
  const dims = getSheetDimensions(ws);
  if (!dims || dims.left === undefined) {
    return ws;
  }
  const start =
    startCol != null
      ? typeof startCol === "string"
        ? colCache.l2n(startCol)
        : startCol
      : dims.left;
  const end =
    endCol != null ? (typeof endCol === "string" ? colCache.l2n(endCol) : endCol) : dims.right;

  for (let c = start; c <= end; c++) {
    _autoFitColumnImpl(ws, c, options);
  }
  return ws;
}

export function autoFitRow(ws: WorksheetData, rowNumber: number): WorksheetData {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  _autoFitRowImpl(ws, rowNumber);
  return ws;
}

export function autoFitRows(ws: WorksheetData, startRow?: number, endRow?: number): WorksheetData {
  checkWorksheetAttached(ws);
  const dims = getSheetDimensions(ws);
  if (!dims || dims.top === undefined) {
    return ws;
  }
  const start = startRow ?? dims.top;
  const end = endRow ?? dims.bottom;

  for (let r = start; r <= end; r++) {
    _autoFitRowImpl(ws, r);
  }
  return ws;
}

/** A row that has never been created has default visibility. */
function _isRowHidden(ws: WorksheetData, rowNumber: number): boolean {
  const row = ws._rows[rowNumber - 1];
  return row ? rowHidden(row) : false;
}

function _isColumnHidden(ws: WorksheetData, colNum: number): boolean {
  return ws._columns[colNum - 1]?.hidden === true;
}

function _anyRowVisible(ws: WorksheetData, top: number, bottom: number): boolean {
  for (let r = top; r <= bottom; r++) {
    if (!_isRowHidden(ws, r)) {
      return true;
    }
  }
  return false;
}

function _anyColumnVisible(ws: WorksheetData, left: number, right: number): boolean {
  for (let c = left; c <= right; c++) {
    if (!_isColumnHidden(ws, c)) {
      return true;
    }
  }
  return false;
}

/**
 * The merged range a master cell heads, or `undefined` for any other cell.
 *
 * A merge's content is stored only in its master, yet it is displayed wherever
 * any part of the merge is. Auto-fit therefore decides a master's visibility
 * over this range rather than its own row or column — which is what keeps a
 * merge whose master is hidden, one inside a collapsed outline group say, from
 * being dropped.
 */
function _mergeRangeOf(ws: WorksheetData, cell: CellData): RangeData | undefined {
  return cellIsMerged(cell) ? ws._merges[cell.address] : undefined;
}

function _autoFitColumnImpl(
  ws: WorksheetData,
  colNum: number,
  options: AutoFitColumnOptions | undefined
): void {
  const mdw = getMaxDigitWidth(); // default font MDW
  const includeHiddenRows = options?.includeHiddenRows === true;

  // Check if this column is under an autofilter
  const hasAutoFilter = _isColumnInAutoFilter(ws, colNum);

  let maxWidthPx = 0;

  // Iterate all rows
  ws._rows.forEach(row => {
    if (!row) {
      return;
    }
    const cell = rowFindCell(row, colNum);
    if (!cell) {
      return;
    }

    // Skip merged cell slaves — the content belongs to the master cell.
    if (cellType(cell) === Enums.ValueType.Merge) {
      return;
    }

    const merge = _mergeRangeOf(ws, cell);
    // A multi-column merge's width should not be attributed to one column.
    if (merge && merge.left !== merge.right) {
      return;
    }
    // Skip content in hidden rows — Excel excludes it from auto-fit
    if (
      !includeHiddenRows &&
      (merge ? !_anyRowVisible(ws, merge.top, merge.bottom) : rowHidden(row))
    ) {
      return;
    }

    // Skip shrinkToFit cells — they adapt to the column, not vice versa
    if (cellAlignment(cell)?.shrinkToFit) {
      return;
    }

    let textWidthPx = getCellTextWidthPx(cellView(cell));

    // Account for indent: each level adds approximately 3 character widths
    const indent = cellAlignment(cell)?.indent;
    if (indent && indent > 0) {
      textWidthPx += indent * 3 * mdw;
    }

    if (textWidthPx > maxWidthPx) {
      maxWidthPx = textWidthPx;
    }
  });

  if (maxWidthPx > 0) {
    const charWidth = calculateAutoFitWidth(maxWidthPx, mdw, hasAutoFilter);
    if (charWidth > 0) {
      const column = getColumn(ws, colNum);
      column.width = charWidth;
      column.customWidth = undefined;
      column.bestFit = true;
    }
  }
}

function _autoFitRowImpl(ws: WorksheetData, rowNumber: number): void {
  const row = ws._rows[rowNumber - 1];
  if (!row) {
    return;
  }

  const mdw = getMaxDigitWidth();
  let maxHeightPt = 0;

  rowEachCell(row, cell => {
    // Skip merged cell slaves
    if (cellType(cell) === Enums.ValueType.Merge) {
      return;
    }
    const merge = _mergeRangeOf(ws, cell);
    // A multi-row merge's height should not be attributed to one row.
    if (merge && merge.top !== merge.bottom) {
      return;
    }
    // Skip content in hidden columns
    const col = cellCol(cell);
    if (merge ? !_anyColumnVisible(ws, merge.left, merge.right) : _isColumnHidden(ws, col)) {
      return;
    }

    const columnWidthPx = _getColumnContentWidthForCell(
      ws,
      cell,
      mdw,
      merge?.left ?? col,
      merge?.right ?? col
    );
    const heightPt = getCellHeightPt(cellView(cell), mdw, columnWidthPx);
    if (heightPt > maxHeightPt) {
      maxHeightPt = heightPt;
    }
  });

  if (maxHeightPt > 0) {
    row.height = Math.ceil(maxHeightPt * 4) / 4; // Round to nearest 0.25pt (Excel precision)
    row.customHeight = true;
  }
}

function _getColumnContentWidthForCell(
  ws: WorksheetData,
  cell: CellData,
  mdw: number,
  left: number,
  right: number
): number | undefined {
  if (!cellAlignment(cell)?.wrapText) {
    return undefined;
  }
  // Text wraps across the visible columns the cell spans, with the cell padding
  // charged once for the whole merge rather than once per column.
  let totalPx = 0;
  for (let c = left; c <= right; c++) {
    if (!_isColumnHidden(ws, c)) {
      totalPx += charWidthToPixel(_columnWidthChars(ws, c), mdw);
    }
  }
  return Math.max(0, totalPx - getPixelPadding(mdw));
}

/** Width of a column in characters, without creating it as a side effect. */
function _columnWidthChars(ws: WorksheetData, colNum: number): number {
  return ws._columns[colNum - 1]?.width ?? ws.properties.defaultColWidth ?? 9;
}

function _isColumnInAutoFilter(ws: WorksheetData, colNum: number): boolean {
  if (!ws.autoFilter) {
    return false;
  }
  if (typeof ws.autoFilter === "string") {
    const range = colCache.decode(ws.autoFilter) as DecodedRange;
    return colNum >= range.left && colNum <= range.right;
  }
  const { from, to } = ws.autoFilter;
  const fromCol =
    typeof from === "string" ? (colCache.decode(from) as { col: number }).col : from.col;
  const toCol = typeof to === "string" ? (colCache.decode(to) as { col: number }).col : to.col;
  return colNum >= fromCol && colNum <= toCol;
}

export function toJSON(ws: WorksheetData, opts: SheetToJsonOptions & { header: 1 }): CellValue[][];

export function toJSON(ws: WorksheetData, opts?: SheetToJsonOptions): Record<string, CellValue>[];

export function toJSON(
  ws: WorksheetData,
  opts?: SheetToJsonOptions
): CellValue[][] | Record<string, CellValue>[] {
  const o = opts || {};

  // Determine range
  let startRow = 1;
  let endRow = getRowCount(ws);
  let startCol = 1;
  let endCol = getColumnCount(ws);

  if (o.range !== undefined) {
    if (typeof o.range === "number") {
      startRow = o.range + 1; // 0-indexed to 1-indexed
    } else if (typeof o.range === "string") {
      const r = decodeRange(o.range);
      startRow = r.s.r + 1;
      endRow = r.e.r + 1;
      startCol = r.s.c + 1;
      endCol = r.e.c + 1;
    }
  }

  if (endRow < startRow || endCol < startCol) {
    return [];
  }

  // `toJSON` is a getter, so it must not materialise the rectangle it walks.
  // A cell that does not exist has no value, and formats to "" — which the
  // callers below already treat as empty.
  const readValue = (row: number, col: number): CellValue => {
    const cell = findCell(ws, row, col);
    if (!cell) {
      return null;
    }
    return cellReadValue(cell);
  };
  const readCell = (row: number, col: number): CellValue => {
    if (o.raw !== false) {
      return readValue(row, col);
    }
    const cell = findCell(ws, row, col);
    return cell ? getCellDisplayText(cellView(cell), o.dateFormat).trim() : null;
  };

  const headerOpt = o.header;

  // header: 1 — return array of arrays
  if (headerOpt === 1) {
    const result: CellValue[][] = [];
    const includeBlank = o.blankRows !== false;

    for (let row = startRow; row <= endRow; row++) {
      const rowData: CellValue[] = [];
      let isEmpty = true;

      for (let col = startCol; col <= endCol; col++) {
        const val = readCell(row, col);

        if (val != null && val !== "") {
          rowData[col - startCol] = val;
          isEmpty = false;
        } else if (o.defaultValue !== undefined) {
          rowData[col - startCol] = o.defaultValue;
        } else {
          rowData[col - startCol] = null;
        }
      }

      if (!isEmpty || includeBlank) {
        result.push(rowData);
      }
    }

    return result;
  }

  // header: "A" — use column letters as keys
  if (headerOpt === "A") {
    const result: Record<string, CellValue>[] = [];
    const includeBlank = o.blankRows === true;

    for (let row = startRow; row <= endRow; row++) {
      const rowData: Record<string, CellValue> = {};
      let isEmpty = true;

      for (let col = startCol; col <= endCol; col++) {
        const val = readCell(row, col);
        const key = encodeCol(col - 1);

        if (val != null && val !== "") {
          rowData[key] = val;
          isEmpty = false;
        } else if (o.defaultValue !== undefined) {
          rowData[key] = o.defaultValue;
        }
      }

      if (!isEmpty || includeBlank) {
        result.push(rowData);
      }
    }

    return result;
  }

  // header: string[] — use provided array as keys
  if (Array.isArray(headerOpt)) {
    const result: Record<string, CellValue>[] = [];
    const includeBlank = o.blankRows === true;

    for (let row = startRow; row <= endRow; row++) {
      const rowData: Record<string, CellValue> = {};
      let isEmpty = true;

      for (let col = startCol; col <= endCol; col++) {
        const colIdx = col - startCol;
        const key = headerOpt[colIdx] ?? `__EMPTY_${colIdx}`;
        const val = readCell(row, col);

        if (val != null && val !== "") {
          rowData[key] = val;
          isEmpty = false;
        } else if (o.defaultValue !== undefined) {
          rowData[key] = o.defaultValue;
        }
      }

      if (!isEmpty || includeBlank) {
        result.push(rowData);
      }
    }

    return result;
  }

  // Default: first row as header, disambiguate duplicates
  const headers: string[] = [];
  const headerCounts: Record<string, number> = {};

  for (let col = startCol; col <= endCol; col++) {
    const val = readValue(startRow, col);
    let header = val != null ? String(val) : `__EMPTY_${col - startCol}`;

    if (headerCounts[header] !== undefined) {
      headerCounts[header]++;
      header = `${header}_${headerCounts[header]}`;
    } else {
      headerCounts[header] = 0;
    }

    headers.push(header);
  }

  const result: Record<string, CellValue>[] = [];
  const dataStartRow = startRow + 1;
  const includeBlank = o.blankRows === true;

  for (let row = dataStartRow; row <= endRow; row++) {
    const rowData: Record<string, CellValue> = {};
    let isEmpty = true;

    for (let col = startCol; col <= endCol; col++) {
      const val = readCell(row, col);
      const key = headers[col - startCol];

      if (val != null && val !== "") {
        rowData[key] = val;
        isEmpty = false;
      } else if (o.defaultValue !== undefined) {
        rowData[key] = o.defaultValue;
      }
    }

    if (!isEmpty || includeBlank) {
      result.push(rowData);
    }
  }

  return result;
}

export function addJSON(
  ws: WorksheetData,
  data: readonly Record<string, CellValue>[],
  opts?: AddJsonOptions
): WorksheetData {
  if (data.length === 0) {
    return ws;
  }

  const o = opts || {};

  // Determine starting position
  let startRow = 1;
  let startCol = 1;
  if (o.origin !== undefined) {
    const resolved = _resolveOrigin(o.origin, getRowCount(ws));
    startRow = resolved.row;
    startCol = resolved.col;
  }

  // Determine headers
  const allKeys = new Set<string>();
  data.forEach(row => Object.keys(row).forEach(k => allKeys.add(k)));
  const headers = o.header ? [...o.header] : [...allKeys];
  if (o.header) {
    allKeys.forEach(k => {
      if (!headers.includes(k)) {
        headers.push(k);
      }
    });
  }

  let rowNum = startRow;

  // Write header row
  if (!o.skipHeader) {
    headers.forEach((h, colIdx) => {
      cellSetValue(getCell(ws, rowNum, startCol + colIdx), h);
    });
    rowNum++;
  }

  // Write data rows
  for (const row of data) {
    headers.forEach((key, colIdx) => {
      const val = row[key];
      if (val === null && o.nullError) {
        cellSetValue(getCell(ws, rowNum, startCol + colIdx), { error: "#NULL!" });
      } else if (val !== undefined && val !== null) {
        cellSetValue(getCell(ws, rowNum, startCol + colIdx), val);
      }
    });
    rowNum++;
  }

  return ws;
}

/**
 * The sheet as an array of rows, **0-based** — row 1 at index `0`, column A at index `0` of each row — with `[]` for a
 * row that does not exist and `null` for a missing cell inside one. Creates nothing.
 */
export function toAOA(ws: WorksheetData): CellValue[][] {
  checkWorksheet(ws);
  const result: CellValue[][] = [];

  // Walk the backing arrays directly. `eachRow`/`rowEachCell` with
  // `includeEmpty` hand back a cell handle for every slot, which means creating
  // one for every hole — and `toAOA` is a getter. Preserve the old output shape:
  // a missing row is `[]`, while a missing cell inside a row is `null`. Index by
  // slot because the hole branch has no row to read a number from; a row always
  // sits at `row.number - 1`, so the two agree.
  for (let i = 0; i < ws._rows.length; i++) {
    const row = ws._rows[i];
    if (!row) {
      result[i] = [];
      continue;
    }
    const rowData: CellValue[] = [];
    for (let j = 0; j < row.cells.length; j++) {
      const cell = row.cells[j];
      rowData[j] = cell ? cellReadValue(cell) : null;
    }
    result[i] = rowData;
  }

  return result;
}

export function addAOA(
  ws: WorksheetData,
  data: readonly (readonly CellValue[])[],
  opts?: AddAoaOptions
): WorksheetData {
  if (data.length === 0) {
    return ws;
  }

  let startRow = 1;
  let startCol = 1;
  if (opts?.origin !== undefined) {
    const resolved = _resolveOrigin(opts.origin, getRowCount(ws));
    startRow = resolved.row;
    startCol = resolved.col;
  }

  data.forEach((row, rowIdx) => {
    if (!row) {
      return;
    }
    row.forEach((val, colIdx) => {
      if (val !== undefined && val !== null) {
        cellSetValue(getCell(ws, startRow + rowIdx, startCol + colIdx), val);
      }
    });
  });

  return ws;
}

/** The used range — the smallest range covering every cell that holds a value. A copy. */
export function getSheetDimensions(ws: WorksheetData): RangeData {
  checkWorksheet(ws);
  const dimensions = rangeCreate();
  ws._rows.forEach(row => {
    if (row) {
      const rowDims = rowDimensions(row);
      if (rowDims) {
        rangeExpand(dimensions, row.number, rowDims.min, row.number, rowDims.max);
      }
    }
  });
  return dimensions;
}

/**
 * The worksheet's **materialised** column records, in column order.
 *
 * A column record exists once it has been declared ({@link setColumns} or
 * `Column.*`) *or* once any cell in that column has been touched — reading or
 * writing `C1` pads the array up to column 3, with columns 1-2 as defaults. So
 * this length tracks the highest column ever addressed, not the number of
 * configured columns, and it is not {@link getColumnCount} either (that counts
 * cells actually present in rows).
 *
 * Read-only: the elements are {@link ColumnView}s, so neither the array nor any
 * column (nor its nested `style`) can be written to. Declare columns through
 * {@link setColumns} and change them through the `Column` namespace — both keep
 * the column numbering, the key registry and the cells' styles in sync. Use
 * {@link getColumnDefinitions} to read the definitions as detached copies.
 */
export function getColumns(ws: WorksheetData): readonly ColumnView[] {
  checkWorksheet(ws);
  return ws._columns;
}

/**
 * The column definitions as plain, detached copies — the inverse of
 * {@link setColumns}. Mirrors {@link getColumns}, so it includes default
 * placeholder entries for columns that exist only because a cell in them was
 * touched; those round-trip as defaults.
 *
 * This is the safe way to clone a sheet's column layout, e.g. to append one:
 * `setColumns(ws, [...getColumnDefinitions(ws), { header: "Error", key: "error" }])`.
 * Copying fields off {@link getColumns} by hand re-implements the `hidden` /
 * `outlineLevel` normalisation and silently drifts when it changes.
 */
export function getColumnDefinitions(ws: WorksheetData): ColumnDefn[] {
  checkWorksheet(ws);
  return ws._columns.map(columnDefn);
}

/** The last **declared** column, as a read-only {@link ColumnView}. */
export function getLastColumn(ws: WorksheetData): ColumnView {
  return getColumn(ws, getColumnCount(ws));
}

/**
 * The width of the **used** range: the largest cell count over all rows.
 *
 * Unrelated to {@link getColumns} / {@link getColumnDefinitions}, which describe
 * the materialised column records — this counts the cells actually present in
 * rows, so `getColumns(ws).length` and `getColumnCount(ws)` are not
 * interchangeable (a column padded into existence by a cell read has a record
 * but no cells).
 * O(rows); {@link getActualColumnCount} additionally ignores gaps.
 */
export function getColumnCount(ws: WorksheetData): number {
  let maxCount = 0;
  eachRow(ws, row => {
    maxCount = Math.max(maxCount, rowCellCount(row));
  });
  return maxCount;
}

export function getActualColumnCount(ws: WorksheetData): number {
  // performance nightmare - for each row, counts all the columns used
  const counts: boolean[] = [];
  let count = 0;
  eachRow(ws, row => {
    rowEachCell(row, (_cell: CellData, col: number) => {
      if (!counts[col]) {
        counts[col] = true;
        count++;
      }
    });
  });
  return count;
}

export function getLastRow(ws: WorksheetData): RowData | undefined {
  // A splice that deletes the last rows leaves holes at the end of `_rows`; the last row is the last one there.
  const n = get_lastRowNumber(ws);
  return n > 0 ? ws._rows[n - 1] : undefined;
}

export function getRowCount(ws: WorksheetData): number {
  return get_lastRowNumber(ws);
}

export function getActualRowCount(ws: WorksheetData): number {
  // counts actual rows that have actual data
  let count = 0;
  eachRow(ws, () => {
    count++;
  });
  return count;
}

/** Whether the sheet has any merged region. */
export function getHasMerges(ws: WorksheetData): boolean {
  checkWorksheet(ws);
  // return true if this._merges has a merge object
  return Object.values(ws._merges).some(Boolean);
}

/** The sheet's merged regions, 1-based and frozen. */
export function getMergedRegions(ws: WorksheetData): ReadonlyArray<{
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}> {
  checkWorksheet(ws);
  return Object.values(ws._merges).map(merge => ({
    top: merge.top,
    left: merge.left,
    bottom: merge.bottom,
    right: merge.right
  }));
}

/**
 * The sheet as its round-trip model — what the writers serialise and {@link setSheetModel} applies.
 *
 * **It shares objects with the sheet**: the cells' value records, the views, the page setup and the other settings are
 * the sheet's own, so editing them edits the sheet without any of the setters' checks. Treat it as read-only, or
 * apply a changed copy with `Worksheet.setModel`.
 */
export function getSheetModel(ws: WorksheetData): WorksheetModel {
  checkWorksheet(ws);
  const model: WorksheetModel = {
    id: ws.id,
    name: getSheetName(ws),
    // **Tab order, which the model dropped.** `orderNo` is allocated across worksheets *and* chartsheets as
    // sheets are added, and it is the only field that records where the author put a tab. Leaving it out of the
    // model made every worksheet compare as "unknown position" while a chartsheet still carried its number — so
    // in `financial-report`, whose chartsheet is added last, the chartsheet sorted *first* and "Board View"
    // appeared as the leading tab in the XLSX. The XLSB writer then produced a third answer of its own.
    //
    // One dropped field, two wrong orders, and both writers looked like the culprit. Neither was.
    orderNo: ws.orderNo,
    // **Notes a streamed sheet retained, because its rows are gone by the time the package is written.**
    //
    // `commentsFromModel` reads notes off `rows[].cells[].comment`, and a streaming worksheet releases `_rows` at
    // `commit()` — so a comment written through the streaming writer reached the model as nothing at all. Measured:
    // streaming XLSX produced `xl/comments1.xml` and `vmlDrawing1.vml`, buffered XLSB produced `xl/comments1.bin` and the
    // VML, and streaming XLSB produced neither — the notes were silently absent, indicator and text alike.
    //
    // Retained rather than streamed because a BIFF12 comments part is written whole at the end, and it is small: one
    // entry per annotated cell, against a row store that may be millions.
    retainedComments: ws._retainedComments,
    // Hyperlinks a streamed sheet retained, for the same reason — see `hyperlinksFromModel`.
    retainedHyperlinks: ws._retainedHyperlinks,
    // Carried so the writer can put the same records back — see `styledBlankRanges` above.
    styledBlankRanges: (ws as { _styledBlankRanges?: readonly StyledBlankRangeModel[] })
      ._styledBlankRanges,
    dataValidations: ws.dataValidations.model,
    properties: ws.properties,
    state: ws.state,
    pageSetup: ws.pageSetup,
    headerFooter: ws.headerFooter,
    rowBreaks: ws.rowBreaks,
    colBreaks: ws.colBreaks,
    views: ws.views,
    autoFilter: ws.autoFilter,
    autoFilterCriteria: ws._autoFilterCriteria,
    sortStateXml: ws._sortStateXml,
    opaqueRels: ws._opaqueRels,
    legacyControls: ws._legacyControls,
    chartsheetPlaceholder: ws._chartsheetPlaceholder,
    xlsbDrawingRelationshipId: ws._xlsbDrawingRelationshipId,
    worksheetNamespaceAttributes: ws._worksheetNamespaceAttributes,
    worksheetMcIgnorable: ws._worksheetMcIgnorable,
    sortStateAutoFilterRef: ws._sortStateAutoFilterRef,
    media: ws._media.map(medium => imageModel(medium)),
    shapes: ws._shapes.map(shape => _resolveShapeModel(ws, shape)),
    sheetProtection: ws.sheetProtection,
    tables: Object.values(ws.tables).map(table => tableGetModel(table)),
    pivotTables: ws.pivotTables,
    conditionalFormattings: ws.conditionalFormattings,
    formControls: ws.formControls.map(fc => fc.model),
    ignoredErrors: ws.ignoredErrors,
    watermark: ws._watermark,
    drawing: ws._drawing,
    charts: ws._charts.map((c): ChartAnchorModel => ({
      chartNumber: c.chartNumber,
      chartExNumber: c.chartExNumber,
      range: {
        tl: {
          nativeCol: c.range.tl.nativeCol,
          nativeColOff: c.range.tl.nativeColOff,
          nativeRow: c.range.tl.nativeRow,
          nativeRowOff: c.range.tl.nativeRowOff
        },
        br: c.range.br
          ? {
              nativeCol: c.range.br.nativeCol,
              nativeColOff: c.range.br.nativeColOff,
              nativeRow: c.range.br.nativeRow,
              nativeRowOff: c.range.br.nativeRowOff
            }
          : undefined,
        editAs: c.range.editAs,
        pos: c.range.pos,
        ext: c.range.ext
      } as ChartAnchorModel["range"]
    })),
    sparklineGroups: ws._sparklineGroups,
    threadedComments: ws.threadedComments
  };

  // =================================================
  // columns
  model.cols = columnToModel(ws._columns);

  // ==========================================================
  // Rows
  const rows: RowModel[] = (model.rows = []);
  const dimensions: RangeData = (model.dimensions = rangeCreate());
  // **`_rows` may be gone.** A streaming worksheet releases it at `commit()` — the rows have been serialised and
  // handed to the ZIP, which is the property that writer exists for — and the workbook-level pass still needs the
  // rest of the model afterwards. So an absent row store means "no rows to describe", not a broken worksheet.
  (ws._rows ?? []).forEach(row => {
    const rowModel = row && rowGetModel(row);
    if (rowModel) {
      rangeExpand(dimensions, rowModel.number, rowModel.min, rowModel.number, rowModel.max);
      rows.push(rowModel);
    }
  });

  // ==========================================================
  // Merges
  model.mergeCells = Object.values(ws._merges).map((merge: RangeData) => rangeRange(merge));

  return model;
}

/**
 * Rename the sheet, and every reference to it — formulas on every sheet, defined names, chart series, pivot sources —
 * with it, as Excel does. A rename used to change the tab and nothing else, so `A!A1*2` elsewhere kept naming a sheet
 * that no longer existed.
 */
export function setSheetName(ws: WorksheetData, name: string | undefined): void {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  const from = ws._name;
  const wb = getSheetWorkbook(ws);
  if (from === undefined || !wb || wb._worksheets[ws.id] !== ws) {
    assignSheetName(ws, name);
    return;
  }
  // Validated before anything is rewritten, so a refused name changes nothing. A name past 31 characters is refused
  // rather than shortened: the caller chose it.
  refuseLongSheetName(name);
  const to = validateSheetName(wb, name ?? `sheet${ws.id}`, ws);
  if (to === from) {
    return;
  }
  withReferenceEdit(wb, { kind: "rename", from, to }, () => {
    ws._name = to;
  });
}

/**
 * Give the sheet its name and change nothing else — for a sheet being created or loaded, where nothing can refer to
 * it under another name yet.
 */
export function assignSheetName(ws: WorksheetData, name: string | undefined): void {
  if (name === undefined) {
    name = `sheet${ws.id}`;
  }

  if (ws._name === name) {
    return;
  }

  // Delegate to the workbook-level validator so both worksheets and
  // chartsheets share a single naming namespace. Previously this
  // setter only cross-checked against other worksheets, allowing a
  // chartsheet named "S" to coexist with a worksheet named "S";
  // Excel itself forbids that collision. `validateSheetName`
  // performs the full type / empty / illegal-char / quote / length
  // / case-insensitive duplicate checks and returns the sanitised
  // name (truncated to 31 chars if needed).
  ws._name = validateSheetName(ws._workbook, name, ws);
}

export function setColumns(ws: WorksheetData, value: readonly ColumnDefn[]): void {
  checkWorksheet(ws);
  checkWorksheetAttached(ws);
  // calculate max header row count
  ws._headerRowCount = value.reduce((pv, cv) => {
    const headerCount = Array.isArray(cv.header) ? cv.header.length : cv.header ? 1 : 0;
    return Math.max(pv, headerCount);
  }, 0);

  // construct Column objects
  let count = 1;
  const columns: ColumnData[] = (ws._columns = []);
  value.forEach(defn => {
    const column = columnCreate(ws, count++, false);
    columns.push(column);
    columnSetDefn(column, defn);
  });
}

// =============================================================================
// Option Types for Data Conversion
// =============================================================================

export interface SheetToJsonOptions {
  /**
   * Control output format:
   * - `1`: Generate an array of arrays
   * - `"A"`: RowData object keys are literal column labels (A, B, C, ...)
   * - `string[]`: Use specified strings as keys in row objects
   * - `undefined`: Read and disambiguate first row as keys
   */
  header?: 1 | "A" | string[];
  /**
   * Override range:
   * - `number`: Use worksheet range but set starting row to the value (0-indexed)
   * - `string`: Use specified range (A1-style bounded range string)
   * - `undefined`: Use worksheet range
   */
  range?: number | string;
  /** Use raw values (true, default) or formatted text strings (false) */
  raw?: boolean;
  /** Default value for empty cells */
  defaultValue?: CellValue;
  /** Include blank rows in output (default: true for AOA, false for objects) */
  blankRows?: boolean;
  /** Override format for date values (only applies when raw: false) */
  dateFormat?: string;
}

export interface AddJsonOptions {
  /** Use specified field order (default: Object.keys from data) */
  header?: string[];
  /** If true, do not include header row in output */
  skipHeader?: boolean;
  /** Starting position: cell address string, {c, r} object, row number (0-indexed), or -1 to append */
  origin?: Origin;
  /** If true, emit #NULL! error cells for null values */
  nullError?: boolean;
}

export interface AddAoaOptions {
  /** Starting position: cell address string, {c, r} object, row number (0-indexed), or -1 to append */
  origin?: Origin;
}

// =============================================================================
// Private Helpers
// =============================================================================

/** Resolve an Origin value to 1-indexed {row, col} for internal use */
function _resolveOrigin(origin: Origin, rowCount: number): { row: number; col: number } {
  if (typeof origin === "string") {
    const addr = decodeCell(origin);
    return { row: addr.r + 1, col: addr.c + 1 };
  }
  if (typeof origin === "number") {
    if (origin === -1) {
      return { row: rowCount + 1, col: 1 };
    }
    return { row: origin + 1, col: 1 };
  }
  return { row: origin.r + 1, col: origin.c + 1 };
}

export type Worksheet = WorksheetData;

// Re-export the worksheet-core container layer so `@excel/worksheet` remains the
// canonical import path for all worksheet operations.
export {
  type WorksheetData,
  type SheetProtection,
  addRow,
  addRows,
  getColumnKey,
  setColumnKey,
  deleteColumnKey,
  getColumn,
  _commitRow,
  findRow,
  findRows,
  getRow,
  getRows,
  eachRow,
  getSheetValues,
  findCell,
  getCell,
  getSheetName,
  getSheetWorkbook,
  getTable,
  getTables,
  rowGetCell,
  rowGetCellEx,
  rowSetValues,
  rowEachCell,
  rowSplice,
  rowCommit,
  rowSetModel,
  columnCreate,
  columnSetDefn,
  columnSetHeader,
  columnSetKey,
  columnEachCell,
  columnValues,
  columnSetValuesAt,
  columnFromModel,
  columnSetNumFmt,
  columnSetFont,
  columnSetAlignment,
  columnSetProtection,
  columnSetBorder,
  columnSetFill,
  freeze,
  split,
  unfreeze,
  panes,
  setAutoFilter,
  autoFilter
} from "@excel/core/worksheet-core";
