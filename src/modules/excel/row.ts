import { Cell, type CellModel, type CellAddress } from "@excel/cell";
import type { Column } from "@excel/column";
import { Enums } from "@excel/enums";
import { ExcelError } from "@excel/errors";
import type {
  Style,
  NumFmt,
  Font,
  Alignment,
  Protection,
  Borders,
  Fill,
  CellValue,
  RowValues,
  RowBreak
} from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import type { Worksheet } from "@excel/worksheet";

// Internal interface for row dimensions
interface RowDimensions {
  min: number;
  max: number;
}

// Internal interface for row model (serialization)
export interface RowModel {
  cells: CellModel[];
  number: number;
  min: number;
  max: number;
  height?: number;
  customHeight?: boolean;
  style: Partial<Style>;
  hidden: boolean;
  outlineLevel: number;
  collapsed: boolean;
  dyDescent?: number;
}

class Row {
  // Type declarations only - no runtime overhead
  declare private _worksheet: Worksheet;
  declare private _number: number;
  declare private _cells: Cell[];
  declare public style: Partial<Style>;
  declare private _hidden?: boolean;
  declare private _outlineLevel?: number;
  declare public height?: number;
  declare public customHeight?: boolean;
  declare public dyDescent?: number;

  constructor(worksheet: Worksheet, number: number) {
    this._worksheet = worksheet;
    this._number = number;
    this._cells = [];
    this.style = {};
    this.outlineLevel = 0;
  }

  /**
   * The row number
   */
  get number(): number {
    return this._number;
  }

  /**
   * The worksheet that contains this row
   */
  get worksheet(): Worksheet {
    return this._worksheet;
  }

  /**
   * Commit a completed row to stream.
   * Inform Streaming Writer that this row (and all rows before it) are complete
   * and ready to write. Has no effect on Worksheet document.
   */
  commit(): void {
    this._worksheet._commitRow(this);
  }

  /**
   * Helps GC by breaking cyclic references
   */
  destroy(): void {
    this._worksheet = undefined!;
    this._cells = undefined!;
    this.style = undefined!;
  }

  findCell(colNumber: number): Cell | undefined {
    return this._cells[colNumber - 1];
  }

  // given {address, row, col}, find or create new cell
  getCellEx(address: CellAddress): Cell {
    let cell = this._cells[address.col - 1];
    if (!cell) {
      const column = this._worksheet.getColumn(address.col);
      cell = new Cell(this, column, address.address);
      this._cells[address.col - 1] = cell;
    }
    return cell;
  }

  /**
   * Get cell by number, column letter or column key
   */
  getCell(col: string | number): Cell {
    let colNum: number;
    if (typeof col === "string") {
      // is it a key?
      const column = this._worksheet.getColumnKey(col);
      if (column) {
        colNum = column.number;
      } else {
        colNum = colCache.l2n(col);
      }
    } else {
      colNum = col;
    }
    return (
      this._cells[colNum - 1] ||
      this.getCellEx({
        address: colCache.encodeAddress(this._number, colNum),
        row: this._number,
        col: colNum
      })
    );
  }

  /**
   * Cut one or more cells (cells to the right are shifted left)
   *
   * Note: this operation will not affect other rows
   */
  splice(start: number, count: number, ...inserts: CellValue[]): void {
    const nKeep = start + count;
    const nExpand = inserts.length - count;
    const nEnd = this._cells.length;
    let i: number;
    let cSrc: Cell | undefined;
    let cDst: Cell | undefined;

    if (nExpand < 0) {
      // remove cells
      for (i = start + inserts.length; i <= nEnd; i++) {
        cDst = this._cells[i - 1];
        cSrc = this._cells[i - nExpand - 1];
        if (cSrc) {
          cDst = this.getCell(i);
          cDst.value = cSrc.value;
          cDst.style = copyStyle(cSrc.style) ?? {};
          cDst.comment = cSrc.comment;
        } else if (cDst) {
          cDst.value = null;
          cDst.style = {};
          cDst.comment = undefined;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        cSrc = this._cells[i - 1];
        if (cSrc) {
          cDst = this.getCell(i + nExpand);
          cDst.value = cSrc.value;
          cDst.style = copyStyle(cSrc.style) ?? {};
          cDst.comment = cSrc.comment;
        } else {
          this._cells[i + nExpand - 1] = undefined!;
        }
      }
    }

    // now add the new values
    for (i = 0; i < inserts.length; i++) {
      cDst = this.getCell(start + i);
      cDst.value = inserts[i];
      cDst.style = {};
      cDst.comment = undefined;
    }
  }

  /**
   * Iterate over all non-null cells in a row
   */
  eachCell(callback: (cell: Cell, colNumber: number) => void): void;
  /**
   * Iterate over all cells in a row (including empty cells)
   */
  eachCell(
    opt: { includeEmpty?: boolean },
    callback: (cell: Cell, colNumber: number) => void
  ): void;
  eachCell(
    optOrCallback: { includeEmpty?: boolean } | ((cell: Cell, colNumber: number) => void),
    maybeCallback?: (cell: Cell, colNumber: number) => void
  ): void {
    let options: { includeEmpty?: boolean } | null = null;
    let callback: (cell: Cell, colNumber: number) => void;
    if (typeof optOrCallback === "function") {
      callback = optOrCallback;
    } else {
      options = optOrCallback;
      callback = maybeCallback!;
    }
    if (options && options.includeEmpty) {
      const n = this._cells.length;
      for (let i = 1; i <= n; i++) {
        callback(this.getCell(i), i);
      }
    } else {
      this._cells.forEach((cell, index) => {
        if (cell && cell.type !== Enums.ValueType.Null) {
          callback(cell, index + 1);
        }
      });
    }
  }

  // ===========================================================================
  // Page Breaks
  addPageBreak(lft?: number, rght?: number): void {
    const ws = this._worksheet;
    const left = Math.max(0, (lft ?? 0) - 1) || 0;
    const right = Math.max(0, (rght ?? 0) - 1) || 16838;
    const pb: RowBreak = {
      id: this._number,
      max: right,
      man: 1
    };
    if (left) {
      pb.min = left;
    }

    ws.rowBreaks.push(pb);
  }

  /**
   * Get a row as a sparse array
   */
  get values(): CellValue[] {
    const values: CellValue[] = [];
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        values[cell.col] = cell.value;
      }
    });
    return values;
  }

  /**
   * Set the values by contiguous or sparse array, or by key'd object literal
   */
  set values(value: RowValues) {
    // this operation is not additive - any prior cells are removed
    this._cells = [];
    if (!value) {
      // empty row
    } else if (value instanceof Array) {
      let offset = 0;
      if (Object.prototype.hasOwnProperty.call(value, "0")) {
        // contiguous array - start at column 1
        offset = 1;
      }
      value.forEach((item, index) => {
        if (item !== undefined) {
          this.getCellEx({
            address: colCache.encodeAddress(this._number, index + offset),
            row: this._number,
            col: index + offset
          }).value = item;
        }
      });
    } else {
      // assume object with column keys
      // Values are typed as `unknown` because RowValues accepts arbitrary objects
      // (e.g. DB entities). Cell.value setter handles unknown values via Value.getType
      // fallback to JSON type, so this cast is safe at runtime.
      this._worksheet.eachColumnKey((column: Column, key: string) => {
        if (value[key] !== undefined) {
          this.getCellEx({
            address: colCache.encodeAddress(this._number, column.number),
            row: this._number,
            col: column.number
          }).value = value[key] as CellValue;
        }
      });
    }
  }

  /**
   * Get row values as a 0-based sparse array (column 1 => index 0).
   *
   * This is useful when you want to compare/stringify values without the
   * leading separator produced by the 1-based sparse array returned by `values`.
   */
  getValues(): CellValue[] {
    const values: CellValue[] = [];
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        values[cell.col - 1] = cell.value;
      }
    });
    return values;
  }

  /**
   * Convenience stringification for row values.
   *
   * Equivalent to `row.getValues().join(separator)`.
   */
  valuesToString(separator = ","): string {
    return this.getValues().join(separator);
  }

  /**
   * Returns true if the row includes at least one cell with a value
   */
  get hasValues(): boolean {
    return this._cells.some(cell => cell && cell.type !== Enums.ValueType.Null);
  }

  /**
   * Number of cells including empty ones
   */
  get cellCount(): number {
    return this._cells.length;
  }

  /**
   * Number of non-empty cells
   */
  get actualCellCount(): number {
    let count = 0;
    this.eachCell(() => {
      count++;
    });
    return count;
  }

  /**
   * Get the min and max column number for the non-null cells in this row or null
   */
  get dimensions(): RowDimensions | null {
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        if (!min || min > cell.col) {
          min = cell.col;
        }
        if (max < cell.col) {
          max = cell.col;
        }
      }
    });
    return min > 0
      ? {
          min,
          max
        }
      : null;
  }

  // =========================================================================
  // styles
  private _applyStyle<K extends keyof Style>(name: K, value: Style[K]): void {
    this.style[name] = value;
    this._cells.forEach(cell => {
      if (cell) {
        // Clone object values so each cell gets an independent copy.
        // Without this, mutating a sub-property (e.g. cell.border.top = ...)
        // would leak to every other cell that received the same reference.
        cell.style[name] =
          typeof value === "object" && value !== null ? structuredClone(value) : value;
      }
    });
  }

  get numFmt(): string | NumFmt | undefined {
    return this.style.numFmt;
  }

  set numFmt(value: string | undefined) {
    if (value !== undefined) {
      this._applyStyle("numFmt", value);
    }
  }

  get font(): Partial<Font> | undefined {
    return this.style.font;
  }

  set font(value: Partial<Font> | undefined) {
    if (value !== undefined) {
      this._applyStyle("font", value);
    }
  }

  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }

  set alignment(value: Partial<Alignment> | undefined) {
    if (value !== undefined) {
      this._applyStyle("alignment", value);
    }
  }

  get protection(): Partial<Protection> | undefined {
    return this.style.protection;
  }

  set protection(value: Partial<Protection> | undefined) {
    if (value !== undefined) {
      this._applyStyle("protection", value);
    }
  }

  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }

  set border(value: Partial<Borders> | undefined) {
    if (value !== undefined) {
      this._applyStyle("border", value);
    }
  }

  get fill(): Fill | undefined {
    return this.style.fill;
  }

  set fill(value: Fill | undefined) {
    if (value !== undefined) {
      this._applyStyle("fill", value);
    }
  }

  get hidden(): boolean {
    return !!this._hidden;
  }

  set hidden(value: boolean) {
    this._hidden = value;
  }

  get outlineLevel(): number {
    return this._outlineLevel ?? 0;
  }

  set outlineLevel(value: number) {
    this._outlineLevel = value;
  }

  get collapsed(): boolean {
    return !!(
      this._outlineLevel && this._outlineLevel >= (this._worksheet.properties.outlineLevelRow ?? 0)
    );
  }

  // =========================================================================
  get model(): RowModel | null {
    const cells: CellModel[] = [];
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell) {
        const cellModel = cell.model;
        if (cellModel) {
          if (!min || min > cell.col) {
            min = cell.col;
          }
          if (max < cell.col) {
            max = cell.col;
          }
          cells.push(cellModel);
        }
      }
    });

    return this.height != null || cells.length
      ? {
          cells,
          number: this.number,
          min,
          max,
          height: this.height,
          customHeight: this.customHeight,
          style: this.style,
          hidden: this.hidden,
          outlineLevel: this.outlineLevel,
          collapsed: this.collapsed,
          dyDescent: this.dyDescent
        }
      : null;
  }

  set model(value: RowModel) {
    if (value.number !== this._number) {
      throw new ExcelError("Invalid row number in model");
    }
    this._cells = [];
    let previousAddress: CellAddress | undefined;
    value.cells.forEach(cellModel => {
      switch (cellModel.type) {
        case Cell.Types.Merge:
          // special case - don't add this types
          break;
        default: {
          let address: CellAddress | undefined;
          if (cellModel.address) {
            address = colCache.decodeAddress(cellModel.address);
          } else if (previousAddress) {
            // This is a <c> element without an r attribute
            // Assume that it's the cell for the next column
            const { row } = previousAddress;
            const col = previousAddress.col + 1;
            address = {
              row,
              col,
              address: colCache.encodeAddress(row, col),
              $col$row: `$${colCache.n2l(col)}$${row}`
            };
          }
          previousAddress = address;
          if (!address) {
            break;
          }
          const cell = this.getCellEx(address);
          cell.model = cellModel;
          break;
        }
      }
    });

    if (value.height != null) {
      this.height = value.height;
    } else {
      delete this.height;
    }

    if (value.customHeight != null) {
      this.customHeight = value.customHeight;
    } else {
      delete this.customHeight;
    }

    this.hidden = value.hidden;
    this.outlineLevel = value.outlineLevel ?? 0;
    this.dyDescent = value.dyDescent;

    this.style = value.style ? structuredClone(value.style) : {};
  }
}

export { Row };
