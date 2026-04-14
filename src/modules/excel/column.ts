import type { Cell, CellValueType } from "@excel/cell";
import { Enums } from "@excel/enums";
import type { Row } from "@excel/row";
import type {
  Style,
  NumFmt,
  Font,
  Alignment,
  Protection,
  Borders,
  Fill,
  CellValue,
  ColBreak
} from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import { isEqual } from "@excel/utils/under-dash";
import type { Worksheet } from "@excel/worksheet";

const DEFAULT_COLUMN_WIDTH = 9;

/** Header value type - can be a single value or array for multi-row headers */
export type ColumnHeaderValue = CellValue | CellValue[];

export interface ColumnDefn {
  /** Column header value(s). Can be string, Date, number, or any CellValue type */
  header?: ColumnHeaderValue;
  key?: string;
  width?: number;
  outlineLevel?: number;
  hidden?: boolean;
  style?: Partial<Style>;
  /** Whether the column width is auto-fitted to content */
  bestFit?: boolean;
}

export interface ColumnModel {
  min: number;
  max: number;
  width?: number;
  style?: Partial<Style>;
  isCustomWidth?: boolean;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  bestFit?: boolean;
}

/**
 * Column defines the column properties for 1 column.
 * This includes header rows, widths, key, (style), etc.
 * Worksheet will condense the columns as appropriate during serialization
 */
class Column {
  declare private _worksheet: Worksheet;
  declare private _number: number;
  declare private _header: ColumnHeaderValue | undefined;
  declare private _key: string | undefined;
  /** The width of the column */
  declare public width?: number;
  declare private _hidden: boolean | undefined;
  declare private _outlineLevel: number | undefined;
  /** Whether the column width is auto-fitted to content */
  declare public bestFit?: boolean;
  /** Styles applied to the column */
  declare public style: Partial<Style>;

  constructor(worksheet: Worksheet, number: number, defn?: ColumnDefn | false) {
    this._worksheet = worksheet;
    this._number = number;
    if (defn !== false) {
      // sometimes defn will follow
      this.defn = defn;
    }
  }

  get number(): number {
    return this._number;
  }

  get worksheet(): Worksheet {
    return this._worksheet;
  }

  /**
   * Column letter key
   */
  get letter(): string {
    return colCache.n2l(this._number);
  }

  get isCustomWidth(): boolean {
    return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
  }

  get defn(): ColumnDefn {
    return {
      header: this._header,
      key: this.key,
      width: this.width,
      style: this.style,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
      bestFit: this.bestFit
    };
  }

  set defn(value: ColumnDefn | undefined) {
    if (value) {
      this.key = value.key;
      this.width = value.width !== undefined ? value.width : DEFAULT_COLUMN_WIDTH;
      this.outlineLevel = value.outlineLevel;
      if (value.style) {
        this.style = (copyStyle(value.style) as Partial<Style>) ?? {};
      } else {
        this.style = {};
      }

      // headers must be set after style
      this.header = value.header;
      this._hidden = !!value.hidden;
      this.bestFit = value.bestFit;
    } else {
      delete this._header;
      delete this._key;
      delete this.width;
      this.style = {};
      this.outlineLevel = 0;
      delete this.bestFit;
    }
  }

  /**
   * Get header values as an array (for multi-row header support)
   */
  get headers(): CellValue[] {
    if (Array.isArray(this._header)) {
      return this._header;
    }
    if (this._header !== undefined) {
      return [this._header];
    }
    return [];
  }

  /**
   * Can be a single value or an array for multi-row headers.
   * Supports any CellValue type including Date, number, string, etc.
   */
  get header(): ColumnHeaderValue | undefined {
    return this._header;
  }

  set header(value: ColumnHeaderValue | undefined) {
    if (value !== undefined) {
      this._header = value;
      this.headers.forEach((cellValue, index) => {
        this._worksheet.getCell(index + 1, this.number).value = cellValue;
      });
    } else {
      this._header = undefined;
    }
  }

  /**
   * The name of the properties associated with this column in each row
   */
  get key(): string | undefined {
    return this._key;
  }

  set key(value: string | undefined) {
    const column = this._key && this._worksheet.getColumnKey(this._key);
    if (column === this) {
      this._worksheet.deleteColumnKey(this._key!);
    }

    this._key = value;
    if (value) {
      this._worksheet.setColumnKey(value, this);
    }
  }

  /**
   * Hides the column
   */
  get hidden(): boolean {
    return !!this._hidden;
  }

  set hidden(value: boolean) {
    this._hidden = value;
  }

  /**
   * Set an outline level for columns
   */
  get outlineLevel(): number {
    return this._outlineLevel ?? 0;
  }

  set outlineLevel(value: number | undefined) {
    this._outlineLevel = value;
  }

  /**
   * Indicate the collapsed state based on outlineLevel
   */
  get collapsed(): boolean {
    return !!(
      this._outlineLevel && this._outlineLevel >= (this._worksheet.properties.outlineLevelCol ?? 0)
    );
  }

  toString(): string {
    return JSON.stringify({
      key: this.key,
      width: this.width,
      headers: this.headers.length ? this.headers : undefined
    });
  }

  equivalentTo(other: Column): boolean {
    return (
      this.width === other.width &&
      this.hidden === other.hidden &&
      this.outlineLevel === other.outlineLevel &&
      isEqual(this.style, other.style)
    );
  }

  equivalentToModel(model: ColumnModel): boolean {
    return (
      this.width === model.width &&
      this.hidden === model.hidden &&
      this.outlineLevel === model.outlineLevel &&
      this.bestFit === model.bestFit &&
      isEqual(this.style, model.style)
    );
  }

  get isDefault(): boolean {
    if (this.isCustomWidth) {
      return false;
    }
    if (this.hidden) {
      return false;
    }
    if (this.outlineLevel) {
      return false;
    }
    if (this.bestFit) {
      return false;
    }
    const s = this.style;
    if (s && (s.font || s.numFmt || s.alignment || s.border || s.fill || s.protection)) {
      return false;
    }
    return true;
  }

  get headerCount(): number {
    return this.headers.length;
  }

  /**
   * Iterate over all current cells in this column
   */
  eachCell(callback: (cell: Cell, rowNumber: number) => void): void;
  /**
   * Iterate over all current cells in this column including empty cells
   */
  eachCell(
    opt: { includeEmpty?: boolean },
    callback: (cell: Cell, rowNumber: number) => void
  ): void;
  eachCell(
    optionsOrCallback: { includeEmpty?: boolean } | ((cell: Cell, rowNumber: number) => void),
    maybeCallback?: (cell: Cell, rowNumber: number) => void
  ): void {
    const colNumber = this.number;
    let options: { includeEmpty?: boolean };
    let callback: (cell: Cell, rowNumber: number) => void;
    if (typeof optionsOrCallback === "function") {
      options = {};
      callback = optionsOrCallback;
    } else {
      options = optionsOrCallback;
      callback = maybeCallback!;
    }
    this._worksheet.eachRow(options, (row: Row, rowNumber: number) => {
      callback(row.getCell(colNumber), rowNumber);
    });
  }

  /**
   * The cell values in the column
   */
  get values(): CellValueType[] {
    const v: CellValueType[] = [];
    this.eachCell((cell, rowNumber) => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        v[rowNumber] = cell.value;
      }
    });
    return v;
  }

  set values(v: CellValueType[]) {
    if (!v) {
      return;
    }
    const colNumber = this.number;
    let offset = 0;
    if (Object.prototype.hasOwnProperty.call(v, "0")) {
      // assume contiguous array, start at row 1
      offset = 1;
    }
    v.forEach((value, index) => {
      this._worksheet.getCell(index + offset, colNumber).value = value;
    });
  }

  // =========================================================================
  // Page Breaks

  /**
   * Add a page break after this column.
   * @param top - Optional top row limit for the page break (1-indexed)
   * @param bottom - Optional bottom row limit for the page break (1-indexed)
   */
  addPageBreak(top?: number, bottom?: number): void {
    const ws = this._worksheet;
    const topRow = Math.max(0, (top ?? 0) - 1) || 0;
    const bottomRow = Math.max(0, (bottom ?? 0) - 1) || 1048575;
    const pb: ColBreak = {
      id: this._number,
      max: bottomRow,
      man: 1
    };
    if (topRow) {
      pb.min = topRow;
    }

    ws.colBreaks.push(pb);
  }

  // =========================================================================
  // styles
  get numFmt(): string | NumFmt | undefined {
    return this.style.numFmt;
  }

  set numFmt(value: string | undefined) {
    this.style.numFmt = value;
    this.eachCell(cell => {
      cell.numFmt = value;
    });
  }

  get font(): Partial<Font> | undefined {
    return this.style.font;
  }

  set font(value: Partial<Font> | undefined) {
    this.style.font = value;
    this.eachCell(cell => {
      cell.font = value ? structuredClone(value) : value;
    });
  }

  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }

  set alignment(value: Partial<Alignment> | undefined) {
    this.style.alignment = value;
    this.eachCell(cell => {
      cell.alignment = value ? structuredClone(value) : value;
    });
  }

  get protection(): Partial<Protection> | undefined {
    return this.style.protection;
  }

  set protection(value: Partial<Protection> | undefined) {
    this.style.protection = value;
    this.eachCell(cell => {
      cell.protection = value ? structuredClone(value) : value;
    });
  }

  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }

  set border(value: Partial<Borders> | undefined) {
    this.style.border = value;
    this.eachCell(cell => {
      cell.border = value ? structuredClone(value) : value;
    });
  }

  get fill(): Fill | undefined {
    return this.style.fill;
  }

  set fill(value: Fill | undefined) {
    this.style.fill = value;
    this.eachCell(cell => {
      cell.fill = value ? structuredClone(value) : value;
    });
  }

  // =============================================================================
  // static functions

  static toModel(columns: Column[]): ColumnModel[] | undefined {
    // Convert array of Column into compressed list cols
    const cols: ColumnModel[] = [];
    let col: ColumnModel | null = null;
    columns.forEach((column, index) => {
      if (column.isDefault) {
        if (col) {
          col = null;
        }
      } else if (!col || !column.equivalentToModel(col)) {
        col = {
          min: index + 1,
          max: index + 1,
          width: column.width !== undefined ? column.width : DEFAULT_COLUMN_WIDTH,
          style: column.style,
          isCustomWidth: column.isCustomWidth,
          hidden: column.hidden,
          outlineLevel: column.outlineLevel,
          collapsed: column.collapsed,
          bestFit: column.bestFit
        };
        cols.push(col);
      } else {
        col.max = index + 1;
      }
    });
    return cols.length ? cols : undefined;
  }

  static fromModel(cols: ColumnModel[]): Column[];
  static fromModel(worksheet: Worksheet, cols: ColumnModel[]): Column[];
  static fromModel(
    worksheetOrCols: Worksheet | ColumnModel[],
    colsMaybe?: ColumnModel[]
  ): Column[] {
    // Streaming readers historically called Column.fromModel(cols) without a Worksheet.
    // Preserve that behavior by accepting the 1-arg form and intentionally using the
    // cols array as the backing "worksheet" object.
    const worksheet: Worksheet = Array.isArray(worksheetOrCols)
      ? (worksheetOrCols as unknown as Worksheet)
      : worksheetOrCols;
    let cols: ColumnModel[] = Array.isArray(worksheetOrCols) ? worksheetOrCols : (colsMaybe ?? []);

    cols = cols ?? [];
    const columns: Column[] = [];
    let count = 1;
    let index = 0;
    /**
     * sort cols by min
     * If it is not sorted, the subsequent column configuration will be overwritten
     * */
    cols = cols.sort((pre, next) => pre.min - next.min);
    while (index < cols.length) {
      const col = cols[index++];
      while (count < col.min) {
        columns.push(new Column(worksheet, count++));
      }
      while (count <= col.max) {
        columns.push(new Column(worksheet, count++, col));
      }
    }
    return columns;
  }
}

export { Column };
