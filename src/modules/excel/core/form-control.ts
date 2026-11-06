import { checkFormCheckbox } from "@excel/core/feature-handles";
import type { Worksheet } from "@excel/core/worksheet";
import { checkWorksheetAttached } from "@excel/core/worksheet-core";
import { ExcelArgumentError } from "@excel/errors";
import { quoteSheetName } from "@excel/utils/address";
import { colCache } from "@excel/utils/col-cache";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import { EMU_PER_PX, EMU_PER_POINT } from "@utils/units";

/**
 * Form Control Checkbox - Legacy checkbox control compatible with Office 2007+ and WPS/LibreOffice
 *
 * Unlike the modern In-Cell Checkbox (which only works in Microsoft 365),
 * Form Control Checkboxes are floating controls that work in virtually all
 * spreadsheet applications.
 */

// ============================================================================
// Constants
// ============================================================================

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
  /**
   * How the checkbox follows the cells under it when rows or columns are inserted, deleted or resized —
   * Excel's *Format Control → Properties → Object positioning*: `"oneCell"` moves it without resizing it,
   * `"twoCell"` moves and resizes it, `"absolute"` leaves it where it is. Defaults to `"oneCell"`, which is
   * what Excel gives a checkbox it inserts, so the checkbox stays beside the cell it is linked to.
   */
  editAs?: "oneCell" | "twoCell" | "absolute";
}

/** Internal model for form control checkbox */
export interface FormCheckboxModel {
  /**
   * VML shape id, unique on the sheet (1025, 1026, … for a new sheet). A write keeps it when no other
   * sheet's VML uses its 1024-id block and renumbers it in the file otherwise, since shape ids are
   * unique across the workbook; this value is not changed.
   */
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
  /**
   * How the checkbox follows the cells under it (see {@link FormCheckboxOptions.editAs}). Absent means
   * `"oneCell"`. Read from the file, so a round trip keeps the positioning its author chose.
   */
  editAs?: "oneCell" | "twoCell" | "absolute";
  /**
   * The control's name (`Check Box 1`, …), as Excel shows it in the Name Box. Absent, a write names it
   * after its shape id. Kept from a read so a round trip does not rename the control.
   */
  name?: string;
  /**
   * The VML shape's CSS-like `style` attribute, kept from a read so a round trip draws the control at the
   * size its author gave it. Absent, a write uses {@link formCheckboxVmlStyle}.
   */
  vmlStyle?: string;
}

/**
 * A legacy form control this library does not model — a button, a list or combo box, an option button, a
 * group box, a spinner, a scroll bar, a label — kept as it was read so that rewriting the sheet does not
 * drop it. The sheet's VML and `<controls>` are regenerated on write, and the control is written back into
 * both: its shape and `<control>` entry verbatim except for the shape id (allocated workbook-wide, like a
 * checkbox's) and the relationship id, and its ctrlProp part byte for byte under a fresh part number.
 */
export interface LegacyControlModel {
  /** VML shape id, as read. A write keeps it when its block is free and renumbers the copy it writes otherwise. */
  shapeId: number;
  /** `<x:ClientData ObjectType>`: `Button`, `Drop`, `List`, `Radio`, `GBox`, `Spin`, `Scroll`, `Label`, … */
  objectType: string;
  /** The control's name, from its `<control>` entry. */
  name?: string;
  /** Corners for the hidden DrawingML bridge anchor: the `<control>` anchor, else the VML one. */
  tl?: FormControlAnchor;
  br?: FormControlAnchor;
  /** The `<v:shape>` element, as read. */
  vmlShape: EchoNode;
  /** The `<v:shapetype>` the shape's `type` names, when the part declared it. */
  vmlShapeType?: EchoNode;
  /** Namespace declarations of the source VML root, for prefixes the shape uses beyond `v`, `o` and `x`. */
  vmlNamespaces?: Record<string, string>;
  /** The sheet's `<control>` element, as read. Absent for a shape no entry names. */
  control?: EchoNode;
  /** The control's `xl/ctrlProps/ctrlPropN.xml`, verbatim. */
  ctrlPropXml?: string;
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
// FormCheckbox — de-classed domain model (data record + flat helpers)
// ============================================================================

/**
 * Plain-data form-control checkbox: the owning worksheet plus its serialized
 * {@link FormCheckboxModel}. Construction parses a range into anchors; the
 * former getter/setters (`checked`/`link`/`text`) and VML helpers become flat
 * functions.
 */
export interface FormCheckboxData {
  worksheet: Worksheet;
  model: FormCheckboxModel;
}

/** Structural guard: is `value` a {@link FormCheckboxData} record? */
export function isFormCheckbox(value: unknown): value is FormCheckboxData {
  return (
    typeof value === "object" &&
    value !== null &&
    "worksheet" in value &&
    "model" in value &&
    typeof (value as FormCheckboxData).model === "object"
  );
}

/**
 * A checkbox link as the absolute reference Excel stores in `fmlaLink`: `c5` → `$C$5`, `'My Sheet'!a1`
 * → `'My Sheet'!$A$1`, `a1:b2` → `$A$1:$B$2`. A sheet prefix is kept (quoted when its name needs it,
 * `''` escapes intact), a leading `=` is dropped, and anything that is not a cell or range — a defined
 * name — is kept as written.
 */
function toAbsoluteRef(ref: string): string {
  let text = ref.trim();
  if (text.startsWith("=")) {
    text = text.slice(1).trim();
  }
  let prefix = "";
  let body = text;
  if (text.startsWith("'")) {
    // A quoted sheet name runs to the first `'` not doubled.
    let i = 1;
    while (i < text.length) {
      if (text[i] === "'") {
        if (text[i + 1] === "'") {
          i += 2;
          continue;
        }
        break;
      }
      i++;
    }
    if (i < text.length && text[i + 1] === "!") {
      prefix = text.slice(0, i + 2);
      body = text.slice(i + 2);
    }
  } else {
    const bang = text.lastIndexOf("!");
    if (bang > 0) {
      prefix = `${quoteSheetName(text.slice(0, bang))}!`;
      body = text.slice(bang + 1);
    }
  }
  const parts = body.split(":");
  if (parts.length > 2) {
    return text;
  }
  const absolute: string[] = [];
  for (const part of parts) {
    const cell = absoluteCell(part);
    if (cell === undefined) {
      return prefix + body;
    }
    absolute.push(cell);
  }
  return prefix + absolute.join(":");
}

/** `c5` / `$C5` / `$C$5` → `$C$5`, or `undefined` when `part` is not a cell inside the grid. */
function absoluteCell(part: string): string | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(part);
  if (!match) {
    return undefined;
  }
  const col = match[1].toUpperCase();
  const row = Number(match[2]);
  if (colCache.l2n(col) > 16384 || row > 1048576) {
    return undefined;
  }
  return `$${col}$${row}`;
}

/** Parse a range input into top-left / bottom-right anchor positions. */
function parseRange(range: FormControlRange): { tl: FormControlAnchor; br: FormControlAnchor } {
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

/** Create a form-control checkbox bound to a worksheet from a range + options. */
export function formCheckboxCreate(
  worksheet: Worksheet,
  range: FormControlRange,
  options?: FormCheckboxOptions
): FormCheckboxData {
  // Parse range to get anchors
  const { tl, br } = parseRange(range);

  // One above every shape id in use (starting from 1025). Deriving it from the
  // count reissued a live id once a checkbox had been removed.
  // A preserved control (a button, a list box) holds its id on the sheet just as a checkbox does.
  const shapeId =
    Math.max(
      (worksheet.formControls ?? []).reduce((max, fc) => Math.max(max, fc.model.shapeId), 1024),
      ...(worksheet._legacyControls ?? []).map(control => control.shapeId)
    ) + 1;

  // Parse link cell reference
  let link: string | undefined;
  if (options?.link) {
    // Ensure absolute reference format
    link = toAbsoluteRef(options.link);
  }

  // Note: ctrlPropId is set later in worksheet-xform.ts prepare() for global uniqueness
  return {
    worksheet,
    model: {
      shapeId,
      ctrlPropId: 0, // Placeholder, set during prepare()
      tl,
      br,
      link,
      checked: options?.checked ? "Checked" : "Unchecked",
      text: options?.text ?? "",
      noThreeD: options?.noThreeD ?? true,
      print: options?.print ?? false,
      editAs: options?.editAs ?? "oneCell"
    }
  };
}

/**
 * Rebuild a checkbox record from a previously-serialised model (e.g.
 * round-tripped via `worksheet.model`). The model is adopted as-is; no range
 * parsing or shape id reassignment is performed.
 */
export function formCheckboxFromModel(
  worksheet: Worksheet,
  model: FormCheckboxModel
): FormCheckboxData {
  // Defensive shallow clone: the caller should not be able to mutate the
  // underlying anchor objects through the original reference.
  return {
    worksheet,
    model: {
      ...model,
      tl: { ...model.tl },
      br: { ...model.br }
    }
  };
}

/** Checkboxes taken off their sheet by `Form.remove`. */
const removedCheckboxes = new WeakSet<FormCheckboxData>();

/** Record that `fc` is no longer on its sheet, so its handle refuses edits from now on. */
export function markCheckboxRemoved(fc: FormCheckboxData): void {
  checkFormCheckbox(fc);
  removedCheckboxes.add(fc);
}

/**
 * Refuse an edit through a checkbox handle whose checkbox was removed, or whose sheet was: the edit changed an object
 * nothing writes any more.
 */
function checkCheckboxEditable(fc: FormCheckboxData): void {
  checkWorksheetAttached(fc.worksheet);
  if (removedCheckboxes.has(fc)) {
    throw new ExcelArgumentError(
      "checkbox",
      "The checkbox was removed from its sheet; it can no longer be edited."
    );
  }
}

/** Whether the checkbox is checked. */
export function formCheckboxChecked(fc: FormCheckboxData): boolean {
  checkFormCheckbox(fc);
  return fc.model.checked === "Checked";
}

/** Set the checkbox checked state. */
export function formCheckboxSetChecked(fc: FormCheckboxData, value: boolean): void {
  checkFormCheckbox(fc);
  checkCheckboxEditable(fc);
  fc.model.checked = value ? "Checked" : "Unchecked";
}

/** The linked cell address (absolute form), if any. */
export function formCheckboxLink(fc: FormCheckboxData): string | undefined {
  checkFormCheckbox(fc);
  return fc.model.link;
}

/** Set the linked cell address (coerced to absolute form). */
export function formCheckboxSetLink(fc: FormCheckboxData, value: string | undefined): void {
  checkFormCheckbox(fc);
  checkCheckboxEditable(fc);
  fc.model.link = value ? toAbsoluteRef(value) : undefined;
}

/** The label text. */
export function formCheckboxText(fc: FormCheckboxData): string {
  checkFormCheckbox(fc);
  return fc.model.text;
}

/** Set the label text. */
export function formCheckboxSetText(fc: FormCheckboxData, value: string): void {
  checkFormCheckbox(fc);
  checkCheckboxEditable(fc);
  fc.model.text = value;
}

/**
 * Convert a checkbox model's anchor to the VML anchor string format:
 * "fromCol, fromColOff, fromRow, fromRowOff, toCol, toColOff, toRow, toRowOff"
 * (VML uses pixels for offsets).
 */
export function formCheckboxVmlAnchor(model: FormCheckboxModel): string {
  const { tl, br } = model;
  const tlColOff = Math.round(tl.colOff / EMU_PER_PX);
  const tlRowOff = Math.round(tl.rowOff / EMU_PER_PX);
  const brColOff = Math.round(br.colOff / EMU_PER_PX);
  const brRowOff = Math.round(br.rowOff / EMU_PER_PX);
  return `${tl.col}, ${tlColOff}, ${tl.row}, ${tlRowOff}, ${br.col}, ${brColOff}, ${br.row}, ${brRowOff}`;
}

/** VML positioning style string for a checkbox model. */
export function formCheckboxVmlStyle(model: FormCheckboxModel): string {
  const marginLeft = Math.round(model.tl.colOff / EMU_PER_POINT);
  const marginTop = Math.round(model.tl.rowOff / EMU_PER_POINT);
  return `position:absolute;margin-left:${marginLeft}pt;margin-top:${marginTop}pt;width:96pt;height:18pt;z-index:1;visibility:visible`;
}

/**
 * A checked state from either form a file uses: the VML/ctrlProp number (`0`/`1`/`2`) or the ctrlProp
 * word (`Checked`/`Unchecked`/`Mixed`). Absent or unrecognised is unchecked, the format's default.
 */
export function formCheckboxStateFromValue(value: string | undefined): CheckboxState {
  switch (value?.trim().toLowerCase()) {
    case "1":
    case "checked":
      return "Checked";
    case "2":
    case "mixed":
      return "Mixed";
    default:
      return "Unchecked";
  }
}

/** Numeric VML checked value for a checkbox model (0, 1, or 2). */
export function formCheckboxVmlCheckedValue(model: FormCheckboxModel): number {
  switch (model.checked) {
    case "Checked":
      return 1;
    case "Mixed":
      return 2;
    default:
      return 0;
  }
}
