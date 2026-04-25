import { colCache } from "@excel/utils/col-cache";
import type { Worksheet } from "@excel/worksheet";

/**
 * Form Control Checkbox - Legacy checkbox control compatible with Office 2007+ and WPS/LibreOffice
 *
 * Unlike the modern In-Cell Checkbox (which only works in Microsoft 365),
 * Form Control Checkboxes are floating controls that work in virtually all
 * spreadsheet applications.
 */

// ============================================================================
// Constants (exported for use by xforms)
// ============================================================================

/** EMU (English Metric Units) to pixels conversion factor at 96 DPI */
export const EMU_PER_PIXEL = 9525;

/** EMU to points conversion factor */
export const EMU_PER_POINT = 12700;

/** Default column offset in EMUs (~15 pixels) */
const DEFAULT_COL_OFF = 142875;

/** Default row offset in EMUs (~3 pixels) */
const DEFAULT_ROW_OFF = 28575;

/** Default end column offset in EMUs (~29 pixels) */
const DEFAULT_END_COL_OFF = 276225;

/** Default end row offset in EMUs (~20 pixels) */
const DEFAULT_END_ROW_OFF = 190500;

// ============================================================================
// Types
// ============================================================================

/** Anchor position for form control placement */
export interface FormControlAnchor {
  /** Column index (0-based) */
  col: number;
  /** Column offset in EMUs (1 pixel ≈ 9525 EMUs at 96 DPI) */
  colOff: number;
  /** Row index (0-based) */
  row: number;
  /** Row offset in EMUs */
  rowOff: number;
}

/** Checkbox state values */
export type CheckboxState = "Checked" | "Unchecked" | "Mixed";

/** Options for adding a form control checkbox */
export interface FormCheckboxOptions {
  /** Cell reference where the checkbox value (TRUE/FALSE) will be stored */
  link?: string;
  /** Initial checked state */
  checked?: boolean;
  /** Label text displayed next to the checkbox */
  text?: string;
  /** Whether to use flat appearance (no 3D effect) */
  noThreeD?: boolean;
  /** Whether to print the checkbox */
  print?: boolean;
}

/** Internal model for form control checkbox */
export interface FormCheckboxModel {
  /** Unique shape ID (e.g., 1025, 1026, ...) */
  shapeId: number;
  /** Control property ID (rId in relationships) */
  ctrlPropId: number;
  /** Relationship id (e.g., rId5) in sheet rels for ctrlProp (set during XLSX prepare) */
  ctrlPropRelId?: string;
  /** Top-left anchor */
  tl: FormControlAnchor;
  /** Bottom-right anchor */
  br: FormControlAnchor;
  /** Cell link (e.g., "$A$1") */
  link?: string;
  /** Checked state */
  checked: CheckboxState;
  /** Label text */
  text: string;
  /** Use flat appearance */
  noThreeD: boolean;
  /** Print control */
  print: boolean;
}

/** Range input for form control - can be a cell reference or position object */
export type FormControlRange =
  | string
  | {
      /** Top-left position */
      tl: { col: number; row: number; colOff?: number; rowOff?: number } | string;
      /** Bottom-right position (optional, defaults to reasonable size) */
      br?: { col: number; row: number; colOff?: number; rowOff?: number } | string;
    }
  | {
      /** Start column (0-based) */
      startCol: number;
      /** Start row (0-based) */
      startRow: number;
      /** End column (0-based) */
      endCol: number;
      /** End row (0-based) */
      endRow: number;
      /** Column offset from start in EMUs */
      startColOff?: number;
      /** Row offset from start in EMUs */
      startRowOff?: number;
      /** Column offset from end in EMUs */
      endColOff?: number;
      /** Row offset from end in EMUs */
      endRowOff?: number;
    };

// ============================================================================
// FormCheckbox Class
// ============================================================================

class FormCheckbox {
  declare public worksheet: Worksheet;
  declare public model: FormCheckboxModel;

  constructor(worksheet: Worksheet, range: FormControlRange, options?: FormCheckboxOptions) {
    this.worksheet = worksheet;

    // Parse range to get anchors
    const { tl, br } = this._parseRange(range);

    // Generate shape ID (starting from 1025)
    const existingCount = worksheet.formControls?.length ?? 0;
    const shapeId = 1025 + existingCount;

    // Parse link cell reference
    let link: string | undefined;
    if (options?.link) {
      // Ensure absolute reference format
      link = this._toAbsoluteRef(options.link);
    }

    // Note: ctrlPropId is set later in worksheet-xform.ts prepare() for global uniqueness
    this.model = {
      shapeId,
      ctrlPropId: 0, // Placeholder, set during prepare()
      tl,
      br,
      link,
      checked: options?.checked ? "Checked" : "Unchecked",
      text: options?.text ?? "",
      noThreeD: options?.noThreeD ?? true,
      print: options?.print ?? false
    };
  }

  /**
   * Rebuild a FormCheckbox from a previously-serialised model (e.g. round-tripped
   * via `worksheet.model`). The model is adopted as-is; no range parsing or shape
   * id reassignment is performed.
   */
  static fromModel(worksheet: Worksheet, model: FormCheckboxModel): FormCheckbox {
    const cb = Object.create(FormCheckbox.prototype) as FormCheckbox;
    cb.worksheet = worksheet;
    // Defensive shallow clone: the caller should not be able to mutate the
    // underlying anchor objects through the original reference.
    cb.model = {
      ...model,
      tl: { ...model.tl },
      br: { ...model.br }
    };
    return cb;
  }

  /**
   * Get the checked state
   */
  get checked(): boolean {
    return this.model.checked === "Checked";
  }

  /**
   * Set the checked state
   */
  set checked(value: boolean) {
    this.model.checked = value ? "Checked" : "Unchecked";
  }

  /**
   * Get the linked cell address
   */
  get link(): string | undefined {
    return this.model.link;
  }

  /**
   * Set the linked cell address
   */
  set link(value: string | undefined) {
    this.model.link = value ? this._toAbsoluteRef(value) : undefined;
  }

  /**
   * Get the label text
   */
  get text(): string {
    return this.model.text;
  }

  /**
   * Set the label text
   */
  set text(value: string) {
    this.model.text = value;
  }

  /**
   * Convert cell reference to absolute format (e.g., "A1" -> "$A$1")
   */
  private _toAbsoluteRef(ref: string): string {
    // If already absolute, return as-is
    if (ref.includes("$")) {
      return ref;
    }
    // Parse and convert
    const addr = colCache.decodeAddress(ref);
    return `$${colCache.n2l(addr.col)}$${addr.row}`;
  }

  /**
   * Parse range input into anchor positions
   */
  private _parseRange(range: FormControlRange): { tl: FormControlAnchor; br: FormControlAnchor } {
    let tl: FormControlAnchor;
    let br: FormControlAnchor;

    if (typeof range === "string") {
      // Parse cell reference like "B2" or range like "B2:D3"
      const isRange = range.includes(":");

      if (isRange) {
        const decoded = colCache.decode(range);

        if ("top" in decoded) {
          // Treat 1-cell ranges (e.g., "J4:J4") as a single cell with default checkbox size.
          if (decoded.left === decoded.right && decoded.top === decoded.bottom) {
            const col = decoded.left - 1;
            const row = decoded.top - 1;
            tl = {
              col,
              colOff: DEFAULT_COL_OFF,
              row,
              rowOff: DEFAULT_ROW_OFF
            };
            br = {
              col: col + 2,
              colOff: DEFAULT_END_COL_OFF,
              row: row + 1,
              rowOff: DEFAULT_END_ROW_OFF
            };
          } else {
            // Regular range
            tl = {
              col: decoded.left - 1, // Convert to 0-based
              colOff: DEFAULT_COL_OFF,
              row: decoded.top - 1,
              rowOff: DEFAULT_ROW_OFF
            };
            br = {
              col: decoded.right - 1,
              colOff: DEFAULT_END_COL_OFF,
              row: decoded.bottom - 1,
              rowOff: DEFAULT_END_ROW_OFF
            };
          }
        } else {
          // Defensive fallback: if the cache returns an address, treat it like a single-cell ref.
          tl = {
            col: decoded.col - 1,
            colOff: DEFAULT_COL_OFF,
            row: decoded.row - 1,
            rowOff: DEFAULT_ROW_OFF
          };
          br = {
            col: decoded.col + 1,
            colOff: DEFAULT_END_COL_OFF,
            row: decoded.row,
            rowOff: DEFAULT_END_ROW_OFF
          };
        }
      } else {
        // Single cell reference - create default size checkbox
        const decoded = colCache.decodeAddress(range);
        tl = {
          col: decoded.col - 1,
          colOff: DEFAULT_COL_OFF,
          row: decoded.row - 1,
          rowOff: DEFAULT_ROW_OFF
        };
        // Default size: about 2 columns wide, 1 row tall
        br = {
          col: decoded.col + 1,
          colOff: DEFAULT_END_COL_OFF,
          row: decoded.row,
          rowOff: DEFAULT_END_ROW_OFF
        };
      }
    } else if ("startCol" in range) {
      // startCol/startRow/endCol/endRow format (0-based)
      tl = {
        col: range.startCol,
        colOff: range.startColOff ?? DEFAULT_COL_OFF,
        row: range.startRow,
        rowOff: range.startRowOff ?? DEFAULT_ROW_OFF
      };
      br = {
        col: range.endCol,
        colOff: range.endColOff ?? DEFAULT_END_COL_OFF,
        row: range.endRow,
        rowOff: range.endRowOff ?? DEFAULT_END_ROW_OFF
      };
    } else {
      // Object format with tl/br
      if (typeof range.tl === "string") {
        const decoded = colCache.decodeAddress(range.tl);
        tl = {
          col: decoded.col - 1,
          colOff: DEFAULT_COL_OFF,
          row: decoded.row - 1,
          rowOff: DEFAULT_ROW_OFF
        };
      } else {
        tl = {
          col: range.tl.col,
          colOff: range.tl.colOff ?? DEFAULT_COL_OFF,
          row: range.tl.row,
          rowOff: range.tl.rowOff ?? DEFAULT_ROW_OFF
        };
      }

      if (range.br) {
        if (typeof range.br === "string") {
          const decoded = colCache.decodeAddress(range.br);
          br = {
            col: decoded.col - 1,
            colOff: DEFAULT_END_COL_OFF,
            row: decoded.row - 1,
            rowOff: DEFAULT_END_ROW_OFF
          };
        } else {
          br = {
            col: range.br.col,
            colOff: range.br.colOff ?? DEFAULT_END_COL_OFF,
            row: range.br.row,
            rowOff: range.br.rowOff ?? DEFAULT_END_ROW_OFF
          };
        }
      } else {
        // Default size
        br = {
          col: tl.col + 2,
          colOff: DEFAULT_END_COL_OFF,
          row: tl.row + 1,
          rowOff: DEFAULT_END_ROW_OFF
        };
      }
    }

    return { tl, br };
  }

  // =========================================================================
  // Instance methods - delegate to static methods
  // =========================================================================

  /**
   * Convert anchor to VML anchor string format
   * Format: "fromCol, fromColOff, fromRow, fromRowOff, toCol, toColOff, toRow, toRowOff"
   * VML uses pixels for offsets
   */
  getVmlAnchor(): string {
    return FormCheckbox.getVmlAnchor(this.model);
  }

  /**
   * Get VML style string for positioning
   */
  getVmlStyle(): string {
    return FormCheckbox.getVmlStyle(this.model);
  }

  /**
   * Get the numeric checked value for VML (0, 1, or 2)
   */
  getVmlCheckedValue(): number {
    return FormCheckbox.getVmlCheckedValue(this.model);
  }

  // =========================================================================
  // Static utility methods - can be used with FormCheckboxModel directly
  // =========================================================================

  /**
   * Convert anchor to VML anchor string format from model
   */
  static getVmlAnchor(model: FormCheckboxModel): string {
    const { tl, br } = model;
    const tlColOff = Math.round(tl.colOff / EMU_PER_PIXEL);
    const tlRowOff = Math.round(tl.rowOff / EMU_PER_PIXEL);
    const brColOff = Math.round(br.colOff / EMU_PER_PIXEL);
    const brRowOff = Math.round(br.rowOff / EMU_PER_PIXEL);
    return `${tl.col}, ${tlColOff}, ${tl.row}, ${tlRowOff}, ${br.col}, ${brColOff}, ${br.row}, ${brRowOff}`;
  }

  /**
   * Get VML style string for positioning from model
   */
  static getVmlStyle(model: FormCheckboxModel): string {
    const marginLeft = Math.round(model.tl.colOff / EMU_PER_POINT);
    const marginTop = Math.round(model.tl.rowOff / EMU_PER_POINT);
    return `position:absolute;margin-left:${marginLeft}pt;margin-top:${marginTop}pt;width:96pt;height:18pt;z-index:1;visibility:visible`;
  }

  /**
   * Get the numeric checked value for VML from model (0, 1, or 2)
   */
  static getVmlCheckedValue(model: FormCheckboxModel): number {
    switch (model.checked) {
      case "Checked":
        return 1;
      case "Mixed":
        return 2;
      default:
        return 0;
    }
  }
}

export { FormCheckbox };
