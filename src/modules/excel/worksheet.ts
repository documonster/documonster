import { colCache, type DecodedRange } from "@excel/utils/col-cache";
import { Range, type RangeInput } from "@excel/range";
import { Row, type RowModel } from "@excel/row";
import { WorksheetNameError, MergeConflictError } from "@excel/errors";
import { Column, type ColumnModel, type ColumnDefn } from "@excel/column";
import type { Cell, FormulaResult, FormulaValueData } from "@excel/cell";
import { Enums } from "@excel/enums";
import { Image, type ImageModel } from "@excel/image";
import { Table, type TableModel } from "@excel/table";
import { DataValidations } from "@excel/data-validations";
import {
  FormCheckbox,
  type FormCheckboxModel,
  type FormCheckboxOptions,
  type FormControlRange
} from "@excel/form-control";
import { Encryptor } from "@excel/utils/encryptor";
import { uint8ArrayToBase64 } from "@utils/utils";
import { makePivotTable, type PivotTable, type PivotTableModel } from "@excel/pivot-table";
import { copyStyle } from "@excel/utils/copy-style";
import { applyMergeBorders, collectMergeBorders } from "@excel/utils/merge-borders";
import { formatCellValue } from "@excel/utils/cell-format";
import { decodeCell, decodeRange, encodeCol, type Origin } from "@excel/utils/address";
import type { Workbook } from "@excel/workbook";
import type {
  AddImageRange,
  AutoFilter,
  CellValue,
  ColBreak,
  ConditionalFormattingOptions,
  DataValidation,
  RowBreak,
  RowValues,
  Style,
  TableProperties,
  WorksheetProperties,
  WorksheetState,
  WorksheetView
} from "@excel/types";

// Type for data validation model - maps address to validation
type DataValidationModel = { [address: string]: DataValidation | undefined };

interface SheetProtection {
  sheet?: boolean;
  objects?: boolean;
  scenarios?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

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

interface PageSetupMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

interface PageSetup {
  margins: PageSetupMargins;
  orientation: string;
  horizontalDpi: number;
  verticalDpi: number;
  fitToPage: boolean;
  pageOrder: string;
  blackAndWhite: boolean;
  draft: boolean;
  cellComments: string;
  errors: string;
  scale: number;
  fitToWidth: number;
  fitToHeight: number;
  paperSize?: number;
  showRowColHeaders: boolean;
  showGridLines: boolean;
  firstPageNumber?: number;
  horizontalCentered: boolean;
  verticalCentered: boolean;
  rowBreaks: RowBreak[];
  printArea?: string;
  printTitlesRow?: string;
  printTitlesColumn?: string;
}

interface HeaderFooter {
  differentFirst: boolean;
  differentOddEven: boolean;
  oddHeader: string | null;
  oddFooter: string | null;
  evenHeader: string | null;
  evenFooter: string | null;
  firstHeader: string | null;
  firstFooter: string | null;
}

interface WorksheetModel {
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
  media: ImageModel[];
  sheetProtection: SheetProtection | null;
  tables: TableModel[];
  pivotTables: PivotTable[];
  conditionalFormattings: ConditionalFormattingOptions[];
  formControls: FormCheckboxModel[];
  cols?: ColumnModel[];
  rows?: RowModel[];
  dimensions?: Range;
  mergeCells?: string[];
  /** Loaded drawing data (for charts, etc.) - preserved for round-trip */
  drawing?: unknown;
}

// Worksheet requirements
//  Operate as sheet inside workbook or standalone
//  Load and Save from file and stream
//  Access/Add/Delete individual cells
//  Manage column widths and row heights

class Worksheet {
  // Type declarations only - no runtime overhead
  declare private _workbook: Workbook;
  declare public id: number;
  declare public orderNo: number;
  declare private _name: string;
  declare public state: WorksheetState;
  declare private _rows: Row[];
  declare private _columns: Column[];
  declare private _keys: { [key: string]: Column };
  declare private _merges: { [key: string]: Range };
  declare public rowBreaks: RowBreak[];
  declare public colBreaks: ColBreak[];
  declare public properties: Partial<WorksheetProperties>;
  declare public pageSetup: PageSetup;
  declare public headerFooter: HeaderFooter;
  declare public dataValidations: DataValidations;
  declare public views: Partial<WorksheetView>[];
  declare public autoFilter: AutoFilter | null;
  declare private _media: Image[];
  declare public sheetProtection: SheetProtection | null;
  declare public tables: { [key: string]: Table };
  declare public pivotTables: PivotTable[];
  declare public conditionalFormattings: ConditionalFormattingOptions[];
  declare public formControls: FormCheckbox[];
  declare private _headerRowCount?: number;
  /** Loaded drawing data (for charts, etc.) - preserved for round-trip */
  declare private _drawing: unknown;

  constructor(options: WorksheetOptions) {
    this._workbook = options.workbook!;

    // in a workbook, each sheet will have a number
    this.id = options.id ?? 0;
    this.orderNo = options.orderNo ?? 0;

    // and a name - use the setter to ensure validation and truncation
    this.name = options.name || `sheet${this.id}`;

    // add a state
    this.state = options.state ?? "visible";

    // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
    // Note: _rows is zero based. Must subtract 1 to go from cell.row to index
    this._rows = [];

    // column definitions
    this._columns = [];

    // column keys (addRow convenience): key ==> this._collumns index
    this._keys = {};

    // keep record of all merges
    this._merges = {};

    // record of all row and column pageBreaks
    this.rowBreaks = [];
    this.colBreaks = [];

    // for tabColor, default row height, outline levels, etc
    this.properties = {
      defaultRowHeight: 15,
      outlineLevelCol: 0,
      outlineLevelRow: 0,
      ...options.properties
    };

    // for all things printing
    this.pageSetup = {
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

    this.headerFooter = {
      differentFirst: false,
      differentOddEven: false,
      oddHeader: null,
      oddFooter: null,
      evenHeader: null,
      evenFooter: null,
      firstHeader: null,
      firstFooter: null,
      ...options.headerFooter
    };

    this.dataValidations = new DataValidations();

    // for freezepanes, split, zoom, gridlines, etc
    this.views = options.views ?? [];

    this.autoFilter = options.autoFilter ?? null;

    // for images, etc
    this._media = [];

    // worksheet protection
    this.sheetProtection = null;

    // for tables
    this.tables = {};

    this.pivotTables = [];

    this.conditionalFormattings = [];

    // for form controls (legacy checkboxes, etc.)
    this.formControls = [];
  }

  get name(): string {
    return this._name;
  }

  set name(name: string | undefined) {
    if (name === undefined) {
      name = `sheet${this.id}`;
    }

    if (this._name === name) {
      return;
    }

    if (typeof name !== "string") {
      throw new WorksheetNameError("The name has to be a string.");
    }

    if (name === "") {
      throw new WorksheetNameError("The name can't be empty.");
    }

    if (name === "History") {
      throw new WorksheetNameError('The name "History" is protected. Please use a different name.');
    }

    // Illegal character in worksheet name: asterisk (*), question mark (?),
    // colon (:), forward slash (/ \), or bracket ([])
    if (/[*?:/\\[\]]/.test(name)) {
      throw new WorksheetNameError(
        `Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`
      );
    }

    if (/(^')|('$)/.test(name)) {
      throw new WorksheetNameError(
        `The first or last character of worksheet name cannot be a single quotation mark: ${name}`
      );
    }

    if (name.length > 31) {
      if (process.env.NODE_ENV !== "production") {
        console.warn(`Worksheet name ${name} exceeds 31 chars. This will be truncated`);
      }
      name = name.substring(0, 31);
    }

    const nameLower = name.toLowerCase();
    if (
      this._workbook.worksheets.find(ws => ws && ws !== this && ws.name.toLowerCase() === nameLower)
    ) {
      throw new WorksheetNameError(`Worksheet name already exists: ${name}`);
    }

    this._name = name;
  }

  /**
   * The workbook that contains this worksheet
   */
  get workbook(): Workbook {
    return this._workbook;
  }

  /**
   * When you're done with this worksheet, call this to remove from workbook
   */
  destroy(): void {
    this._workbook.removeWorksheetEx(this);
  }

  /**
   * Get the bounding range of the cells in this worksheet
   */
  get dimensions(): Range {
    const dimensions = new Range();
    this._rows.forEach(row => {
      if (row) {
        const rowDims = row.dimensions;
        if (rowDims) {
          dimensions.expand(row.number, rowDims.min, row.number, rowDims.max);
        }
      }
    });
    return dimensions;
  }

  // =========================================================================
  // Columns

  /**
   * Get the current columns array
   */
  get columns(): Column[] {
    return this._columns;
  }

  /**
   * Add column headers and define column keys and widths.
   *
   * Note: these column structures are a workbook-building convenience only,
   * apart from the column width, they will not be fully persisted.
   */
  set columns(value: ColumnDefn[]) {
    // calculate max header row count
    this._headerRowCount = value.reduce((pv, cv) => {
      const headerCount = Array.isArray(cv.header) ? cv.header.length : cv.header ? 1 : 0;
      return Math.max(pv, headerCount);
    }, 0);

    // construct Column objects
    let count = 1;
    const columns: Column[] = (this._columns = []);
    value.forEach(defn => {
      const column = new Column(this, count++, false);
      columns.push(column);
      column.defn = defn;
    });
  }

  getColumnKey(key: string): Column | undefined {
    return this._keys[key];
  }

  setColumnKey(key: string, value: Column): void {
    this._keys[key] = value;
  }

  deleteColumnKey(key: string): void {
    delete this._keys[key];
  }

  eachColumnKey(f: (column: Column, key: string) => void): void {
    Object.keys(this._keys).forEach(key => f(this._keys[key], key));
  }

  /**
   * Access an individual column by key, letter and 1-based column number
   */
  getColumn(c: string | number): Column {
    let colNum: number;
    if (typeof c === "string") {
      // if it matches a key'd column, return that
      const col = this._keys[c];
      if (col) {
        return col;
      }

      // otherwise, assume letter
      colNum = colCache.l2n(c);
    } else {
      colNum = c;
    }
    if (colNum > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= colNum) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[colNum - 1];
  }

  /**
   * Cut one or more columns (columns to the right are shifted left)
   * and optionally insert more
   *
   * If column properties have been defined, they will be cut or moved accordingly
   *
   * Known limitation: If a splice causes any merged cells to move, the results may be unpredictable
   *
   * Also: If the worksheet has more rows than values in the column inserts,
   * the rows will still be shifted as if the values existed
   */
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void {
    // Before splicing cells, release all cell-level merge references so that
    // row.splice copies plain values instead of merge proxies.
    // _spliceMerges (called later) will rebuild cell-level refs at new coordinates.
    for (const merge of Object.values(this._merges)) {
      for (let r = merge.top; r <= merge.bottom; r++) {
        for (let c = merge.left; c <= merge.right; c++) {
          const cell = this.findCell(r, c);
          if (cell && cell.type === Enums.ValueType.Merge) {
            cell.unmerge();
          }
        }
      }
    }

    const rows = this._rows;
    const nRows = rows.length;
    if (inserts.length > 0) {
      // must iterate over all rows whether they exist yet or not
      for (let i = 0; i < nRows; i++) {
        const insertValues = inserts.map(insert => insert[i] ?? null);
        const row = this.getRow(i + 1);
        row.splice(start, count, ...insertValues);
      }
    } else {
      // nothing to insert, so just splice all rows
      this._rows.forEach(r => {
        if (r) {
          r.splice(start, count);
        }
      });
    }

    // splice column definitions
    const nExpand = inserts.length - count;
    const nKeep = start + count;
    const nEnd = this._columns.length;
    if (nExpand < 0) {
      for (let i = start + inserts.length; i <= nEnd; i++) {
        this.getColumn(i).defn = this.getColumn(i - nExpand).defn;
      }
    } else if (nExpand > 0) {
      for (let i = nEnd; i >= nKeep; i--) {
        this.getColumn(i + nExpand).defn = this.getColumn(i).defn;
      }
    }
    for (let i = start; i < start + inserts.length; i++) {
      this.getColumn(i).defn = undefined;
    }

    // account for defined names
    this.workbook.definedNames.spliceColumns(this.name, start, count, inserts.length);

    // account for images
    if (nExpand !== 0) {
      for (const image of this._media) {
        if (image.type === "image" && image.range) {
          const { tl, br } = image.range;
          if (tl && tl.nativeCol >= start - 1) {
            tl.nativeCol = Math.max(0, tl.nativeCol + nExpand);
          }
          if (br && br.nativeCol >= start - 1) {
            br.nativeCol = Math.max(0, br.nativeCol + nExpand);
          }
        }
      }
    }

    // account for merges
    this._spliceMerges("col", start, count, inserts.length);
  }

  /**
   * Get the last column in a worksheet
   */
  get lastColumn(): Column {
    return this.getColumn(this.columnCount);
  }

  /**
   * The total column size of the document. Equal to the maximum cell count from all of the rows
   */
  get columnCount(): number {
    let maxCount = 0;
    this.eachRow(row => {
      maxCount = Math.max(maxCount, row.cellCount);
    });
    return maxCount;
  }

  /**
   * A count of the number of columns that have values
   */
  get actualColumnCount(): number {
    // performance nightmare - for each row, counts all the columns used
    const counts: boolean[] = [];
    let count = 0;
    this.eachRow(row => {
      row.eachCell(({ col }: { col: number }) => {
        if (!counts[col]) {
          counts[col] = true;
          count++;
        }
      });
    });
    return count;
  }

  // =========================================================================
  // Rows

  /** @internal */
  _commitRow(row: Row): void {
    // nop - allows streaming reader to fill a document
  }

  private get _lastRowNumber(): number {
    // need to cope with results of splice
    const rows = this._rows;
    let n = rows.length;
    while (n > 0 && rows[n - 1] === undefined) {
      n--;
    }
    return n;
  }

  private get _nextRow(): number {
    return this._lastRowNumber + 1;
  }

  /**
   * Get the last editable row in a worksheet (or undefined if there are none)
   */
  get lastRow(): Row | undefined {
    if (this._rows.length) {
      return this._rows[this._rows.length - 1];
    }
    return undefined;
  }

  /**
   * Tries to find and return row for row number, else undefined
   *
   * @param r - The 1-indexed row number
   */
  findRow(r: number): Row | undefined {
    return this._rows[r - 1];
  }

  /**
   * Tries to find and return rows for row number start and length, else undefined
   *
   * @param start - The 1-indexed starting row number
   * @param length - The length of the expected array
   */
  findRows(start: number, length: number): (Row | undefined)[] {
    return this._rows.slice(start - 1, start - 1 + length);
  }

  /**
   * The total row size of the document. Equal to the row number of the last row that has values.
   */
  get rowCount(): number {
    return this._lastRowNumber;
  }

  /**
   * A count of the number of rows that have values. If a mid-document row is empty, it will not be included in the count.
   */
  get actualRowCount(): number {
    // counts actual rows that have actual data
    let count = 0;
    this.eachRow(() => {
      count++;
    });
    return count;
  }

  // get a row by row number.
  getRow(r: number): Row {
    let row = this._rows[r - 1];
    if (!row) {
      row = this._rows[r - 1] = new Row(this, r);
    }
    return row;
  }

  // get multiple rows by row number.
  getRows(start: number, length: number): Row[] | undefined {
    if (length < 1) {
      return undefined;
    }
    const rows: Row[] = [];
    for (let i = start; i < start + length; i++) {
      rows.push(this.getRow(i));
    }
    return rows;
  }

  addRow(value: RowValues, style: string = "n"): Row {
    const rowNo = this._nextRow;
    const row = this.getRow(rowNo);
    row.values = value;
    this._setStyleOption(rowNo, style[0] === "i" ? style : "n");
    return row;
  }

  addRows(value: RowValues[], style: string = "n"): Row[] {
    const rows: Row[] = [];
    value.forEach(row => {
      rows.push(this.addRow(row, style));
    });
    return rows;
  }

  insertRow(pos: number, value: RowValues, style: string = "n"): Row {
    this.spliceRows(pos, 0, value);
    this._setStyleOption(pos, style);
    return this.getRow(pos);
  }

  insertRows(pos: number, values: RowValues[], style: string = "n"): Row[] | undefined {
    this.spliceRows(pos, 0, ...values);
    if (style !== "n") {
      // copy over the styles
      for (let i = 0; i < values.length; i++) {
        if (style[0] === "o" && this.findRow(values.length + pos + i) !== undefined) {
          this._copyStyle(values.length + pos + i, pos + i, style[1] === "+");
        } else if (style[0] === "i" && this.findRow(pos - 1) !== undefined) {
          this._copyStyle(pos - 1, pos + i, style[1] === "+");
        }
      }
    }
    return this.getRows(pos, values.length);
  }

  // set row at position to same style as of either pervious row (option 'i') or next row (option 'o')
  private _setStyleOption(pos: number, style: string = "n"): void {
    if (style[0] === "o" && this.findRow(pos + 1) !== undefined) {
      this._copyStyle(pos + 1, pos, style[1] === "+");
    } else if (style[0] === "i" && this.findRow(pos - 1) !== undefined) {
      this._copyStyle(pos - 1, pos, style[1] === "+");
    }
  }

  private _copyStyle(src: number, dest: number, styleEmpty: boolean = false): void {
    const rSrc = this.getRow(src);
    const rDst = this.getRow(dest);
    rDst.style = (copyStyle(rSrc.style) as Partial<Style>) ?? {};
    rSrc.eachCell({ includeEmpty: styleEmpty }, (cell: Cell, colNumber: number) => {
      rDst.getCell(colNumber).style = (copyStyle(cell.style) as Partial<Style>) ?? {};
    });
    rDst.height = rSrc.height;
  }

  /**
   * Duplicate rows and insert new rows
   */
  duplicateRow(rowNum: number, count: number, insert: boolean = false): void {
    // create count duplicates of rowNum
    // either inserting new or overwriting existing rows

    const rSrc = this.getRow(rowNum);
    const inserts = Array.from<RowValues>({ length: count }).fill(rSrc.values);

    // Collect single-row merges from the source row before splicing
    // (only merges where top == bottom == rowNum, i.e. horizontal merges within one row)
    const srcMerges: Range[] = [];
    for (const merge of Object.values(this._merges)) {
      if (merge.top === rowNum && merge.bottom === rowNum) {
        srcMerges.push(merge);
      }
    }

    // Collect images anchored to the source row before splicing
    // (images whose top-left anchor is on the source row)
    const srcImages: Image[] = [];
    const srcRow0 = rowNum - 1; // 0-based source row
    for (const image of this._media) {
      if (image.type === "image" && image.range) {
        if (image.range.tl.nativeRow === srcRow0) {
          srcImages.push(image);
        }
      }
    }

    this.spliceRows(rowNum + 1, insert ? 0 : count, ...inserts);

    // now copy styles...
    for (let i = 0; i < count; i++) {
      const rDst = this._rows[rowNum + i];
      rDst.style = copyStyle(rSrc.style) ?? {};
      rDst.height = rSrc.height;
      rSrc.eachCell({ includeEmpty: true }, (cell: Cell, colNumber: number) => {
        rDst.getCell(colNumber).style = copyStyle(cell.style) ?? {};
      });
    }

    // Duplicate single-row merges from source row into each new row
    if (srcMerges.length > 0) {
      for (let i = 0; i < count; i++) {
        const dstRow = rowNum + 1 + i;

        // In overwrite mode, clear any existing merges in the target row
        if (!insert) {
          const toRemove: string[] = [];
          for (const [key, merge] of Object.entries(this._merges)) {
            if (merge.top <= dstRow && merge.bottom >= dstRow) {
              toRemove.push(key);
            }
          }
          for (const key of toRemove) {
            this._unMergeMaster(this.getCell(key));
          }
        }

        for (const srcMerge of srcMerges) {
          this.mergeCellsWithoutStyle(dstRow, srcMerge.left, dstRow, srcMerge.right);
        }
      }
    }

    // Duplicate images from source row into each new row.
    // In overwrite mode, first remove any images anchored to the target rows
    // so they don't coexist with the clones (mirrors merge cleanup above).
    if (!insert) {
      const dstStart0 = rowNum; // first target row, 0-based (1-based rowNum + 1 → 0-based rowNum)
      const dstEnd0 = rowNum + count - 1; // last target row, 0-based
      this._media = this._media.filter(image => {
        if (image.type === "image" && image.range) {
          const row0 = image.range.tl.nativeRow;
          return row0 < dstStart0 || row0 > dstEnd0;
        }
        return true;
      });
    }

    for (let i = 0; i < count; i++) {
      const rowDelta = i + 1; // offset from source row to target row
      for (const srcImage of srcImages) {
        const cloned = srcImage.clone();
        cloned.range!.tl.nativeRow = srcRow0 + rowDelta;
        if (cloned.range!.br) {
          const brDelta = srcImage.range!.br!.nativeRow - srcRow0;
          cloned.range!.br.nativeRow = srcRow0 + rowDelta + brDelta;
        }
        this._media.push(cloned);
      }
    }
  }

  /**
   * Cut one or more rows (rows below are shifted up)
   * and optionally insert more
   *
   * Known limitation: If a splice causes any merged cells to move, the results may be unpredictable
   */
  spliceRows(start: number, count: number, ...inserts: RowValues[]): void {
    // same problem as row.splice, except worse.

    // Before splicing rows, release all cell-level merge references so that
    // row value copies work on plain values instead of merge proxies.
    // _spliceMerges (called later) will rebuild cell-level refs at new coordinates.
    for (const merge of Object.values(this._merges)) {
      for (let r = merge.top; r <= merge.bottom; r++) {
        for (let c = merge.left; c <= merge.right; c++) {
          const cell = this.findCell(r, c);
          if (cell && cell.type === Enums.ValueType.Merge) {
            cell.unmerge();
          }
        }
      }
    }

    const nKeep = start + count;
    const nInserts = inserts.length;
    const nExpand = nInserts - count;
    const nEnd = this._rows.length;
    let i: number;
    let rSrc: Row | undefined;
    if (nExpand < 0) {
      // remove rows
      if (start === nEnd) {
        this._rows[nEnd - 1] = undefined!;
      }
      for (i = nKeep; i <= nEnd; i++) {
        rSrc = this._rows[i - 1];
        if (rSrc) {
          const rDst = this.getRow(i + nExpand);
          rDst.values = rSrc.values;
          rDst.style = (copyStyle(rSrc.style) as Partial<Style>) ?? {};
          rDst.height = rSrc.height;
          rSrc.eachCell({ includeEmpty: true }, (cell: Cell, colNumber: number) => {
            rDst.getCell(colNumber).style = (copyStyle(cell.style) as Partial<Style>) ?? {};
          });
          this._rows[i - 1] = undefined!;
        } else {
          this._rows[i + nExpand - 1] = undefined!;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        rSrc = this._rows[i - 1];
        if (rSrc) {
          const rDst = this.getRow(i + nExpand);
          rDst.values = rSrc.values;
          rDst.style = (copyStyle(rSrc.style) as Partial<Style>) ?? {};
          rDst.height = rSrc.height;
          rSrc.eachCell({ includeEmpty: true }, (cell: Cell, colNumber: number) => {
            rDst.getCell(colNumber).style = (copyStyle(cell.style) as Partial<Style>) ?? {};
          });
        } else {
          this._rows[i + nExpand - 1] = undefined!;
        }
      }
    }

    // now copy over the new values
    for (i = 0; i < nInserts; i++) {
      const rDst = this.getRow(start + i);
      rDst.style = {};
      rDst.values = inserts[i];
    }

    // account for defined names
    this.workbook.definedNames.spliceRows(this.name, start, count, nInserts);

    // account for images
    if (nExpand !== 0) {
      for (const image of this._media) {
        if (image.type === "image" && image.range) {
          const { tl, br } = image.range;
          if (tl && tl.nativeRow >= start - 1) {
            tl.nativeRow = Math.max(0, tl.nativeRow + nExpand);
          }
          if (br && br.nativeRow >= start - 1) {
            br.nativeRow = Math.max(0, br.nativeRow + nExpand);
          }
        }
      }
    }

    // account for merges
    this._spliceMerges("row", start, count, nInserts);
  }

  /**
   * Iterate over all rows that have values in a worksheet
   */
  eachRow(callback: (row: Row, rowNumber: number) => void): void;
  /**
   * Iterate over all rows (including empty rows) in a worksheet
   */
  eachRow(opt: { includeEmpty?: boolean }, callback: (row: Row, rowNumber: number) => void): void;
  eachRow(
    optOrCallback: { includeEmpty?: boolean } | ((row: Row, rowNumber: number) => void),
    maybeCallback?: (row: Row, rowNumber: number) => void
  ): void {
    let options: { includeEmpty?: boolean } | undefined;
    let callback: (row: Row, rowNumber: number) => void;
    if (typeof optOrCallback === "function") {
      callback = optOrCallback;
    } else {
      options = optOrCallback;
      callback = maybeCallback!;
    }
    if (options && options.includeEmpty) {
      const n = this._rows.length;
      for (let i = 1; i <= n; i++) {
        callback(this.getRow(i), i);
      }
    } else {
      this._rows.forEach(row => {
        if (row && row.hasValues) {
          callback(row, row.number);
        }
      });
    }
  }

  /**
   * Return all rows as sparse array
   */
  getSheetValues(): CellValue[][] {
    const rows: CellValue[][] = [];
    this._rows.forEach(row => {
      if (row) {
        rows[row.number] = row.values;
      }
    });
    return rows;
  }

  // =========================================================================
  // Cells

  /**
   * Returns the cell at [r,c] or address given by r. If not found, return undefined
   */
  findCell(r: number | string, c?: number): Cell | undefined {
    const address = colCache.getAddress(r, c);
    const row = this._rows[address.row - 1];
    return row ? row.findCell(address.col) : undefined;
  }

  /**
   * Get or create cell at [r,c] or address given by r
   */
  getCell(r: number | string, c?: number): Cell {
    const address = colCache.getAddress(r, c);
    const row = this.getRow(address.row);
    return row.getCellEx(address);
  }

  // =========================================================================
  // Merge

  /**
   * Merge cells, either:
   *
   * tlbr string, e.g. `'A4:B5'`
   *
   * tl string, br string, e.g. `'G10', 'H11'`
   *
   * t, l, b, r numbers, e.g. `10,11,12,13`
   */
  mergeCells(...cells: RangeInput[]): void {
    const dimensions = new Range(cells);
    this._mergeCellsInternal(dimensions);
  }

  mergeCellsWithoutStyle(...cells: RangeInput[]): void {
    const dimensions = new Range(cells);
    this._mergeCellsInternal(dimensions, true);
  }

  private _mergeCellsInternal(dimensions: Range, ignoreStyle?: boolean): void {
    // check cells aren't already merged
    Object.values(this._merges).forEach((merge: Range) => {
      if (merge.intersects(dimensions)) {
        throw new MergeConflictError();
      }
    });

    const { top, left, bottom, right } = dimensions;

    // Collect perimeter borders BEFORE merge overwrites slave styles
    const collected = ignoreStyle
      ? undefined
      : collectMergeBorders(top, left, bottom, right, (r, c) => this.findCell(r, c) as any);

    // Apply merge — slave cells inherit the master's full style
    const master = this.getCell(dimensions.top, dimensions.left);
    for (let i = top; i <= bottom; i++) {
      for (let j = left; j <= right; j++) {
        if (i > top || j > left) {
          this.getCell(i, j).merge(master, ignoreStyle);
        }
      }
    }

    // Reconstruct position-aware borders (like Excel):
    // outer borders survive, inner borders are cleared.
    if (collected) {
      applyMergeBorders(top, left, bottom, right, collected, (r, c) => this.getCell(r, c) as any);
    }

    // index merge
    this._merges[master.address] = dimensions;
  }

  private _unMergeMaster(master: Cell): void {
    // master is always top left of a rectangle
    const merge = this._merges[master.address];
    if (merge) {
      for (let i = merge.top; i <= merge.bottom; i++) {
        for (let j = merge.left; j <= merge.right; j++) {
          this.getCell(i, j).unmerge();
        }
      }
      delete this._merges[master.address];
    }
  }

  /**
   * Update _merges dictionary and cell-level merge references after a row or column splice.
   */
  private _spliceMerges(axis: "row" | "col", start: number, count: number, nInserts: number): void {
    const nExpand = nInserts - count;
    if (nExpand === 0 && count === 0) {
      return;
    }
    const nKeep = start + count;
    const isRow = axis === "row";

    const newMerges: Record<string, Range> = {};

    for (const merge of Object.values(this._merges)) {
      const { top, left, bottom, right } = merge.model;
      // For row axis: lo=top, hi=bottom. For col axis: lo=left, hi=right.
      const lo = isRow ? top : left;
      const hi = isRow ? bottom : right;

      if (nExpand <= 0 && count > 0) {
        // Deleting rows/columns
        const deleteEnd = nKeep - 1;
        if (lo > deleteEnd) {
          // Entirely after deleted range — shift
          const newRange = isRow
            ? new Range(top + nExpand, left, bottom + nExpand, right)
            : new Range(top, left + nExpand, bottom, right + nExpand);
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
          const newRange = new Range(newTop, newLeft, newBottom, newRight);
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
            ? new Range(top + nExpand, left, bottom + nExpand, right)
            : new Range(top, left + nExpand, bottom, right + nExpand);
          newMerges[colCache.encodeAddress(newRange.top, newRange.left)] = newRange;
        } else if (hi < nKeep) {
          // Entirely before splice — unchanged
          newMerges[colCache.encodeAddress(top, left)] = merge;
        } else {
          // Spans splice boundary — stretch
          if (isRow) {
            merge.model.bottom = bottom + nExpand;
          } else {
            merge.model.right = right + nExpand;
          }
          newMerges[colCache.encodeAddress(top, left)] = merge;
        }
      }
    }

    this._merges = newMerges;

    // Rebuild cell-level merge references for all merges.
    // Pre-unmerge in spliceRows/spliceColumns clears all cell refs,
    // so we must rebuild every merge, not just moved/resized ones.
    for (const m of Object.values(newMerges)) {
      const master = this.getCell(m.top, m.left);
      for (let r = m.top; r <= m.bottom; r++) {
        for (let c = m.left; c <= m.right; c++) {
          if (r > m.top || c > m.left) {
            this.getCell(r, c).merge(master, true);
          }
        }
      }
    }
  }

  get hasMerges(): boolean {
    // return true if this._merges has a merge object
    return Object.values(this._merges).some(Boolean);
  }

  /**
   * Scan the range and if any cell is part of a merge, un-merge the group.
   * Note this function can affect multiple merges and merge-blocks are
   * atomic - either they're all merged or all un-merged.
   */
  unMergeCells(...cells: RangeInput[]): void {
    const dimensions = new Range(cells);

    // find any cells in that range and unmerge them
    for (let i = dimensions.top; i <= dimensions.bottom; i++) {
      for (let j = dimensions.left; j <= dimensions.right; j++) {
        const cell = this.findCell(i, j);
        if (cell) {
          if (cell.type === Enums.ValueType.Merge) {
            // this cell merges to another master
            this._unMergeMaster(cell.master);
          } else if (this._merges[cell.address]) {
            // this cell is a master
            this._unMergeMaster(cell);
          }
        }
      }
    }
  }

  // ===========================================================================
  // Shared/Array Formula
  fillFormula(
    range: string,
    formula: string,
    results?:
      | FormulaResult[][]
      | FormulaResult[]
      | ((row: number, col: number) => FormulaResult | undefined),
    shareType: string = "shared"
  ): void {
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
          const cell = this.getCell(r, c);
          const formulaValue: FormulaValueData = {
            shareType,
            formula,
            ref: range,
            result: getResult(r, c)
          };
          cell.value = formulaValue as CellValue;
          first = false;
        } else {
          this.getCell(r, c).value = isShared
            ? {
                sharedFormula: masterAddress,
                result: getResult(r, c)
              }
            : getResult(r, c);
        }
      }
    }
  }

  // =========================================================================
  // Images

  /**
   * Using the image id from `Workbook.addImage`,
   * embed an image within the worksheet to cover a range
   */
  addImage(imageId: string | number, range: AddImageRange): void {
    const model = {
      type: "image",
      imageId: String(imageId),
      range
    };
    this._media.push(new Image(this, model));
  }

  getImages(): Image[] {
    return this._media.filter(m => m.type === "image");
  }

  /**
   * Using the image id from `Workbook.addImage`, set the background to the worksheet
   */
  addBackgroundImage(imageId: string | number): void {
    const model = {
      type: "background",
      imageId: String(imageId)
    };
    this._media.push(new Image(this, model));
  }

  getBackgroundImageId(): string | undefined {
    const image = this._media.find(m => m.type === "background");
    return image && image.imageId;
  }

  // =========================================================================
  // Form Controls (Legacy Checkboxes)

  /**
   * Add a form control checkbox to the worksheet.
   *
   * Form control checkboxes are the legacy style that work in Office 2007+,
   * WPS Office, LibreOffice, and other spreadsheet applications.
   *
   * Unlike modern in-cell checkboxes (which only work in Microsoft 365),
   * form control checkboxes are floating controls positioned over cells.
   *
   * @param range - Cell reference (e.g., "B2") or range (e.g., "B2:D3") for positioning
   * @param options - Checkbox options
   * @returns The created FormCheckbox instance
   *
   * @example
   * // Simple checkbox at B2
   * ws.addFormCheckbox("B2");
   *
   * // Checkbox with label and linked cell
   * ws.addFormCheckbox("B2:D3", {
   *   text: "Accept terms",
   *   link: "A2",
   *   checked: false
   * });
   */
  addFormCheckbox(range: FormControlRange, options?: FormCheckboxOptions): FormCheckbox {
    const checkbox = new FormCheckbox(this, range, options);
    this.formControls.push(checkbox);
    return checkbox;
  }

  /**
   * Get all form control checkboxes in the worksheet
   */
  getFormCheckboxes(): FormCheckbox[] {
    return this.formControls;
  }

  // =========================================================================
  // Worksheet Protection

  /**
   * Protect the worksheet with optional password and options
   */
  async protect(password?: string, options?: Partial<SheetProtection>): Promise<void> {
    this.sheetProtection = {
      sheet: true
    };
    if (options && "spinCount" in options) {
      // force spinCount to be integer >= 0
      options.spinCount = Number.isFinite(options.spinCount)
        ? Math.round(Math.max(0, options.spinCount!))
        : 100000;
    }
    if (password) {
      this.sheetProtection.algorithmName = "SHA-512";
      this.sheetProtection.saltValue = uint8ArrayToBase64(Encryptor.randomBytes(16));
      this.sheetProtection.spinCount =
        options && "spinCount" in options ? options.spinCount : 100000; // allow user specified spinCount
      this.sheetProtection.hashValue = await Encryptor.convertPasswordToHash(
        password,
        "SHA-512",
        this.sheetProtection.saltValue!,
        this.sheetProtection.spinCount!
      );
    }
    if (options) {
      this.sheetProtection = Object.assign(this.sheetProtection, options);
      if (!password && "spinCount" in options) {
        delete this.sheetProtection.spinCount;
      }
    }
  }

  unprotect(): void {
    this.sheetProtection = null;
  }

  // =========================================================================
  // Tables

  /**
   * Add a new table and return a reference to it
   */
  addTable(model: TableProperties): Table {
    const table = new Table(this, model);
    // Use table.name (sanitized by Table.validate()) as the key
    this.tables[table.name] = table;
    return table;
  }

  /**
   * Fetch table by name
   */
  getTable(name: string): Table {
    return this.tables[name];
  }

  /**
   * Delete table by name
   */
  removeTable(name: string): void {
    delete this.tables[name];
  }

  /**
   * Fetch all tables in the worksheet
   */
  getTables(): Table[] {
    return Object.values(this.tables);
  }

  // =========================================================================
  // Pivot Tables
  addPivotTable(model: PivotTableModel): PivotTable {
    const pivotTable = makePivotTable(this, model);

    this.pivotTables.push(pivotTable);
    this.workbook.pivotTables.push(pivotTable);

    return pivotTable;
  }

  // ===========================================================================
  // Conditional Formatting

  /**
   * Add conditional formatting rules
   */
  addConditionalFormatting(cf: ConditionalFormattingOptions): void {
    this.conditionalFormattings.push(cf);
  }

  /**
   * Delete conditional formatting rules
   */
  removeConditionalFormatting(
    filter:
      | number
      | ((
          value: ConditionalFormattingOptions,
          index: number,
          array: ConditionalFormattingOptions[]
        ) => boolean)
  ): void {
    if (typeof filter === "number") {
      this.conditionalFormattings.splice(filter, 1);
    } else if (filter instanceof Function) {
      this.conditionalFormattings = this.conditionalFormattings.filter(filter);
    } else {
      this.conditionalFormattings = [];
    }
  }

  // ===========================================================================
  // Model

  get model(): WorksheetModel {
    const model: WorksheetModel = {
      id: this.id,
      name: this.name,
      dataValidations: this.dataValidations.model,
      properties: this.properties,
      state: this.state,
      pageSetup: this.pageSetup,
      headerFooter: this.headerFooter,
      rowBreaks: this.rowBreaks,
      colBreaks: this.colBreaks,
      views: this.views,
      autoFilter: this.autoFilter,
      media: this._media.map(medium => medium.model),
      sheetProtection: this.sheetProtection,
      tables: Object.values(this.tables).map(table => table.model),
      pivotTables: this.pivotTables,
      conditionalFormattings: this.conditionalFormattings,
      formControls: this.formControls.map(fc => fc.model),
      drawing: this._drawing
    };

    // =================================================
    // columns
    model.cols = Column.toModel(this.columns);

    // ==========================================================
    // Rows
    const rows: RowModel[] = (model.rows = []);
    const dimensions: Range = (model.dimensions = new Range());
    this._rows.forEach(row => {
      const rowModel = row && row.model;
      if (rowModel) {
        dimensions.expand(rowModel.number, rowModel.min, rowModel.number, rowModel.max);
        rows.push(rowModel);
      }
    });

    // ==========================================================
    // Merges
    model.mergeCells = Object.values(this._merges).map((merge: Range) => merge.range);

    return model;
  }

  private _parseRows(model: WorksheetModel): void {
    this._rows = [];
    if (model.rows) {
      model.rows.forEach(rowModel => {
        const row = new Row(this, rowModel.number);
        this._rows[row.number - 1] = row;
        row.model = rowModel;
      });
    }
  }

  private _parseMergeCells(model: WorksheetModel): void {
    if (model.mergeCells) {
      model.mergeCells.forEach((merge: string) => {
        // Do not merge styles when importing an Excel file
        // since each cell may have different styles intentionally.
        this.mergeCellsWithoutStyle(merge);
      });
    }
  }

  set model(value: WorksheetModel) {
    this.name = value.name;
    this.state = value.state;
    this._columns = Column.fromModel(this, value.cols ?? []);
    this._parseRows(value);

    this._parseMergeCells(value);
    this.dataValidations = new DataValidations(value.dataValidations);
    this.properties = value.properties;
    this.pageSetup = value.pageSetup;
    this.headerFooter = value.headerFooter;
    this.rowBreaks = value.rowBreaks ?? [];
    this.colBreaks = value.colBreaks ?? [];
    this.views = value.views;
    this.autoFilter = value.autoFilter;
    this._media = value.media.map(medium => new Image(this, medium));
    this.sheetProtection = value.sheetProtection;
    this.tables = value.tables.reduce((tables: { [key: string]: Table }, table: TableModel) => {
      const t = new Table(this, table);
      t.model = table;
      tables[table.name] = t;
      return tables;
    }, {});
    this.pivotTables = value.pivotTables;
    this.conditionalFormattings = value.conditionalFormattings;
    // Form controls are currently write-only (not parsed from XLSX)
    this.formControls = [];
    // Preserve loaded drawing data (charts, etc.)
    this._drawing = value.drawing;
  }

  // ===========================================================================
  // Data Conversion — JSON
  // ===========================================================================

  /**
   * Convert worksheet data to a JSON array.
   *
   * @example
   * // Default: first row as headers, returns array of objects
   * const data = ws.toJSON();
   * // => [{name: "Alice", age: 30}, {name: "Bob", age: 25}]
   *
   * @example
   * // Array of arrays
   * const aoa = ws.toJSON({ header: 1 });
   * // => [["name", "age"], ["Alice", 30], ["Bob", 25]]
   *
   * @example
   * // Column letters as keys
   * const cols = ws.toJSON({ header: "A" });
   * // => [{A: "name", B: "age"}, {A: "Alice", B: 30}]
   */
  toJSON(opts: SheetToJSONOptions & { header: 1 }): CellValue[][];
  toJSON(opts?: SheetToJSONOptions): Record<string, CellValue>[];
  toJSON(opts?: SheetToJSONOptions): CellValue[][] | Record<string, CellValue>[] {
    const o = opts || {};

    // Determine range
    let startRow = 1;
    let endRow = this.rowCount;
    let startCol = 1;
    let endCol = this.columnCount;

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

    const headerOpt = o.header;

    // header: 1 — return array of arrays
    if (headerOpt === 1) {
      const result: CellValue[][] = [];
      const includeBlank = o.blankRows !== false;

      for (let row = startRow; row <= endRow; row++) {
        const rowData: CellValue[] = [];
        let isEmpty = true;

        for (let col = startCol; col <= endCol; col++) {
          const cell = this.getCell(row, col);
          const val = o.raw === false ? _getCellDisplayText(cell, o.dateFormat).trim() : cell.value;

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
          const cell = this.getCell(row, col);
          const val = o.raw === false ? _getCellDisplayText(cell, o.dateFormat).trim() : cell.value;
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
          const cell = this.getCell(row, col);
          const val = o.raw === false ? _getCellDisplayText(cell, o.dateFormat).trim() : cell.value;

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
      const cell = this.getCell(startRow, col);
      const val = cell.value;
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
        const cell = this.getCell(row, col);
        const val = o.raw === false ? _getCellDisplayText(cell, o.dateFormat).trim() : cell.value;
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

  /**
   * Add data from a JSON array to this worksheet.
   * Each object's keys become column headers (written in the first row unless skipHeader is set).
   *
   * @example
   * ws.addJSON([{name: "Alice", age: 30}, {name: "Bob", age: 25}]);
   *
   * @returns this (for chaining)
   */
  addJSON(data: Record<string, CellValue>[], opts?: AddJSONOptions): this {
    if (data.length === 0) {
      return this;
    }

    const o = opts || {};

    // Determine starting position
    let startRow = 1;
    let startCol = 1;
    if (o.origin !== undefined) {
      const resolved = _resolveOrigin(o.origin, this.rowCount);
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
        this.getCell(rowNum, startCol + colIdx).value = h;
      });
      rowNum++;
    }

    // Write data rows
    for (const row of data) {
      headers.forEach((key, colIdx) => {
        const val = row[key];
        if (val === null && o.nullError) {
          this.getCell(rowNum, startCol + colIdx).value = { error: "#NULL!" };
        } else if (val !== undefined && val !== null) {
          this.getCell(rowNum, startCol + colIdx).value = val;
        }
      });
      rowNum++;
    }

    return this;
  }

  // ===========================================================================
  // Data Conversion — Array of Arrays
  // ===========================================================================

  /**
   * Convert worksheet data to an array of arrays.
   *
   * @example
   * const aoa = ws.toAOA();
   * // => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]
   */
  toAOA(): CellValue[][] {
    const result: CellValue[][] = [];

    this.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const rowData: CellValue[] = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        rowData[colNumber - 1] = cell.value;
      });
      result[rowNumber - 1] = rowData;
    });

    return result;
  }

  /**
   * Add data from an array of arrays to this worksheet.
   *
   * @example
   * ws.addAOA([["Name", "Age"], ["Alice", 30], ["Bob", 25]]);
   *
   * @returns this (for chaining)
   */
  addAOA(data: CellValue[][], opts?: AddAOAOptions): this {
    if (data.length === 0) {
      return this;
    }

    let startRow = 1;
    let startCol = 1;
    if (opts?.origin !== undefined) {
      const resolved = _resolveOrigin(opts.origin, this.rowCount);
      startRow = resolved.row;
      startCol = resolved.col;
    }

    data.forEach((row, rowIdx) => {
      if (!row) {
        return;
      }
      row.forEach((val, colIdx) => {
        if (val !== undefined && val !== null) {
          this.getCell(startRow + rowIdx, startCol + colIdx).value = val;
        }
      });
    });

    return this;
  }
}

// =============================================================================
// Option Types for Data Conversion
// =============================================================================

export interface SheetToJSONOptions {
  /**
   * Control output format:
   * - `1`: Generate an array of arrays
   * - `"A"`: Row object keys are literal column labels (A, B, C, ...)
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

export interface AddJSONOptions {
  /** Use specified field order (default: Object.keys from data) */
  header?: string[];
  /** If true, do not include header row in output */
  skipHeader?: boolean;
  /** Starting position: cell address string, {c, r} object, row number (0-indexed), or -1 to append */
  origin?: Origin;
  /** If true, emit #NULL! error cells for null values */
  nullError?: boolean;
}

export interface AddAOAOptions {
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

/** Get formatted display text for a cell value */
function _getCellDisplayText(cell: Cell, dateFormat?: string): string {
  const value = cell.value;
  const numFmt = cell.numFmt;
  const fmt = typeof numFmt === "string" ? numFmt : (numFmt?.formatCode ?? "General");

  if (value == null) {
    return "";
  }

  if (
    value instanceof Date ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "string"
  ) {
    return formatCellValue(value, fmt, dateFormat);
  }

  // Formula type — use the result value
  if (typeof value === "object" && "formula" in value) {
    const result = value.result;
    if (result == null) {
      return "";
    }
    if (
      result instanceof Date ||
      typeof result === "number" ||
      typeof result === "boolean" ||
      typeof result === "string"
    ) {
      return formatCellValue(result, fmt, dateFormat);
    }
  }

  // Fallback to cell.text for other types (rich text, hyperlink, error, etc.)
  return cell.text;
}

export { Worksheet, type WorksheetModel };
