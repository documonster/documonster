import type { Cell } from "@excel/cell";
import type { Column } from "@excel/column";
import { Enums } from "@excel/enums";
import { MergeConflictError } from "@excel/errors";
import { type Dimensions, type RangeInput, Range } from "@excel/range";
import type { Row } from "@excel/row";
import type { AutoFilter, WorksheetProperties } from "@excel/types";
import { formatCellValue } from "@excel/utils/cell-format";
import { colCache, type DecodedRange } from "@excel/utils/col-cache";
import { applyMergeBorders, collectMergeBorders } from "@excel/utils/merge-borders";

import {
  measureTextWidthPx,
  measureRichTextWidthPx,
  calculateAutoFitHeight,
  calculateRichTextAutoFitHeight,
  getMaxDigitWidth,
  getColumnContentWidthPx
} from "@excel/utils/text-metrics";

// Abstract woksheet class
// implements the same functionality for Worksheet and WorksheetWriter
export abstract class AbstractWorksheet {
  declare protected _rows: Row[];
  declare protected _columns: Column[];
  declare protected _merges: { [key: string]: Range };

  abstract dimensions: Dimensions;
  declare properties: Partial<WorksheetProperties>;
  declare autoFilter: AutoFilter | null;

  abstract getColumn(c: string | number): Column;
  abstract getCell(r: string | number, c?: number): Cell;
  abstract findCell(r: string | number, c?: number): Cell | undefined;
  abstract findRow(rowNumber: number): Row | undefined

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

  // ===========================================================================
  // Auto-Fit row height

  /**
   * Auto-fit a single row's height to its content.
   *
   * @param rowNumber - Row number (1-based)
   * @returns The worksheet (for chaining)
   */
  autoFitRow(rowNumber: number): this {
    this._autoFitRowImpl(rowNumber);
    return this;
  }

  /**
   * Auto-fit all rows (or a range of rows) to their content.
   *
   * @param startRow - Start row (1-based). Defaults to first row.
   * @param endRow - End row (1-based). Defaults to last row.
   * @returns The worksheet (for chaining)
   */
  autoFitRows(startRow?: number, endRow?: number): this {
    const dims = this.dimensions;
    if (!dims || dims.top === undefined) {
      return this;
    }
    const start = startRow ?? dims.top;
    const end = endRow ?? dims.bottom;

    for (let r = start; r <= end; r++) {
      this._autoFitRowImpl(r);
    }
    return this;
  }

  /** Get formatted display text for a cell value */
  protected _getCellDisplayText(cell: Cell, dateFormat?: string): string {
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


  /** @internal Implementation of row auto-fit */
  private _autoFitRowImpl(rowNumber: number): void {
    const row = this.findRow(rowNumber);

    if (!row) {
      return;
    }

    const mdw = getMaxDigitWidth();
    let maxHeightPt = 0;

    row.eachCell(cell => {
      // Skip merged cell slaves
      if (cell.type === Enums.ValueType.Merge) {
        return;
      }
      // Skip multi-row merged masters
      if (cell.isMerged) {
        const mergeRange = this._merges[cell.address];
        if (mergeRange && mergeRange.top !== mergeRange.bottom) {
          return;
        }
      }
      // Skip cells in hidden columns
      const col = this._columns[cell.col - 1];
      if (col?.hidden) {
        return;
      }

      const heightPt = this._getCellHeightPt(cell, mdw);
      if (heightPt > maxHeightPt) {
        maxHeightPt = heightPt;
      }
    });

    if (maxHeightPt > 0) {
      row.height = Math.ceil(maxHeightPt * 4) / 4; // Round to nearest 0.25pt (Excel precision)
      row.customHeight = true;
    }
  }

  /**
   * @internal Get the pixel width of a cell's display text.
   * Handles all cell value types: string, number (formatted), date (formatted),
   * boolean, formula result, rich text, hyperlink, error.
   */
  protected _getCellTextWidthPx(cell: Cell): number {
    const cellType = cell.effectiveType;
    const font = cell.font;

    // Rich text: measure per-run with individual fonts
    if (cellType === Enums.ValueType.RichText) {
      const value = cell.value;
      if (value && typeof value === "object" && "richText" in value) {
        return measureRichTextWidthPx(value.richText, font);
      }
    }

    // Get the display text (applies number formatting)
    const displayText = this._getCellDisplayText(cell);
    if (!displayText) {
      return 0;
    }

    return measureTextWidthPx(displayText, font);
  }

  /**
   * @internal Get the height in points a cell needs.
   * Considers wrapText alignment, indent, and explicit newlines.
   */
  private _getCellHeightPt(cell: Cell, mdw: number): number {
    const font = cell.font;
    const alignment = cell.alignment;
    const cellType = cell.effectiveType;

    // Rich text
    if (cellType === Enums.ValueType.RichText) {
      const value = cell.value;
      if (value && typeof value === "object" && "richText" in value) {
        const columnWidthPx = this._getColumnContentWidthForCell(cell, mdw);
        return calculateRichTextAutoFitHeight(value.richText, font, alignment, columnWidthPx);
      }
    }

    const displayText = this._getCellDisplayText(cell);
    if (!displayText) {
      return 0;
    }

    const columnWidthPx = alignment?.wrapText
      ? this._getColumnContentWidthForCell(cell, mdw)
      : undefined;

    return calculateAutoFitHeight(displayText, font, alignment, columnWidthPx);
  }

  /**
   * @internal Get the content width of the column a cell belongs to, in pixels.
   * Uses the explicit column width if set, otherwise falls back to the worksheet
   * default or the Excel default (9 character units ≈ 64px).
   */
  private _getColumnContentWidthForCell(cell: Cell, mdw: number): number | undefined {
    if (!cell.alignment?.wrapText) {
      return undefined;
    }
    // Try to get explicit column width; avoid creating a column as side effect
    const col = this._columns[cell.col - 1];
    const colWidth = col?.width ?? this.properties.defaultColWidth ?? 9;
    return getColumnContentWidthPx(colWidth, mdw);
  }

  /** @internal Check if a column falls within the autoFilter range */
  protected _isColumnInAutoFilter(colNum: number): boolean {
    if (!this.autoFilter) {
      return false;
    }
    if (typeof this.autoFilter === "string") {
      const range = colCache.decode(this.autoFilter) as DecodedRange;
      return colNum >= range.left && colNum <= range.right;
    }
    const { from, to } = this.autoFilter;
    const fromCol =
      typeof from === "string" ? (colCache.decode(from) as { col: number }).col : from.col;
    const toCol = typeof to === "string" ? (colCache.decode(to) as { col: number }).col : to.col;
    return colNum >= fromCol && colNum <= toCol;
  }
}

