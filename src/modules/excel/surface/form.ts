import { checkFormCheckbox } from "@excel/core/feature-handles";
import type {
  FormCheckboxData,
  FormCheckboxModel,
  FormCheckboxOptions,
  FormControlRange
} from "@excel/core/form-control";
import {
  formCheckboxChecked,
  formCheckboxCreate,
  formCheckboxFromModel,
  formCheckboxLink,
  formCheckboxSetChecked,
  formCheckboxSetLink,
  formCheckboxSetText,
  formCheckboxText,
  formCheckboxVmlAnchor,
  formCheckboxVmlCheckedValue,
  formCheckboxVmlStyle,
  isFormCheckbox
} from "@excel/core/form-control";
import { definedNameProblem } from "@excel/core/table-name";
import {
  addFormCheckbox,
  getFormCheckboxes,
  listFormCheckboxes,
  removeFormCheckbox
} from "@excel/core/worksheet";
/**
 * `Form` namespace surface — legacy form controls.
 *
 * Worksheet-level:
 *   `Form.addCheckbox(ws, range, opts)`, `Form.list(ws)`,
 *   `Form.remove(ws, cb)`.
 * Checkbox-handle operations:
 *   `Form.isCheckbox(v)`, `Form.isChecked(cb)`, `Form.setChecked(cb, true)`,
 *   `Form.getLink(cb)`, `Form.setText(cb, "Label")`, …
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { InvalidAddressError } from "@excel/errors";
import { checkListTarget } from "@excel/surface/argument-checks";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks } from "@excel/utils/arguments";

/** The sheet's checkboxes, in the order they were added — a copy of the list. */
export function list(ws: WorksheetData): FormCheckboxData[] {
  checkWorksheet(ws);
  return listFormCheckboxes(ws);
}

/**
 * Remove a form control — the handle {@link list} returned, or its index in that list — the `remove` paired with
 * `list` over the sheet's controls, whatever their kind. Returns `true` if one was removed.
 */
export function remove(ws: WorksheetData, control: FormCheckboxData | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Form.remove", "control", control, "a form control");
  return removeFormCheckbox(ws, control);
}

/** Whether `value` is a checkbox handle. */
export { isFormCheckbox as isCheckbox } from "@excel/core/form-control";

const GRID = { col: 16383, row: 1048575 };

/** A 0-based grid position with optional EMU offsets. */
function checkPosition(
  check: ArgumentChecks,
  field: string,
  position: unknown,
  keys: { col: string; row: string; colOff: string; rowOff: string }
): void {
  if (!check.record(field, position)) {
    return;
  }
  const at = (key: string) => `${field}.${key}`;
  check.required(at(keys.col), position[keys.col]);
  check.number(at(keys.col), position[keys.col], { min: 0, max: GRID.col, integer: true });
  check.required(at(keys.row), position[keys.row]);
  check.number(at(keys.row), position[keys.row], { min: 0, max: GRID.row, integer: true });
  check.number(at(keys.colOff), position[keys.colOff], { min: 0, integer: true });
  check.number(at(keys.rowOff), position[keys.rowOff], { min: 0, integer: true });
}

const CORNER = { col: "col", row: "row", colOff: "colOff", rowOff: "rowOff" };

function checkCorner(check: ArgumentChecks, field: string, corner: unknown): void {
  if (typeof corner === "string") {
    check.reference(field, corner);
  } else {
    check.required(field, corner);
    checkPosition(check, field, corner, CORNER);
  }
}

function checkRange(check: ArgumentChecks, range: unknown): void {
  check.required("range", range);
  if (typeof range === "string") {
    check.reference("range", range, { range: true });
    return;
  }
  if (!check.record("range", range)) {
    return;
  }
  if ("startCol" in range) {
    const start = {
      col: "startCol",
      row: "startRow",
      colOff: "startColOff",
      rowOff: "startRowOff"
    };
    const end = { col: "endCol", row: "endRow", colOff: "endColOff", rowOff: "endRowOff" };
    checkPosition(check, "range", range, start);
    checkPosition(check, "range", range, end);
    return;
  }
  checkCorner(check, "range.tl", range.tl);
  if (range.br !== undefined) {
    checkCorner(check, "range.br", range.br);
  }
}

/**
 * A linked cell: a cell or range (`B2`, `$B$2`, `Sheet2!B2`, a leading `=` allowed) or a defined name, either one
 * optionally after a sheet name. Anything else was written into the control as a formula Excel cannot resolve —
 * `"not a ref !!"` became `'not a ref !'!`.
 */
function checkLink(check: ArgumentChecks, field: string, link: unknown): void {
  check.string(field, link);
  if (typeof link !== "string") {
    return;
  }
  let text = link.trim();
  if (text.startsWith("=")) {
    text = text.slice(1).trim();
  }
  // A sheet qualifier (`Sheet2!`, `'It''s'!`) may come first; what follows is a cell, a range or a defined name.
  const bang = text.lastIndexOf("!");
  const body = bang >= 0 ? text.slice(bang + 1) : text;
  const sheet = bang >= 0 ? text.slice(0, bang) : undefined;
  if (sheet !== undefined && (sheet === "" || sheet === "''")) {
    check.fail(field, "a cell, a range or a defined name, optionally after a sheet name", link);
  }
  if (definedNameProblem(body) === undefined) {
    return;
  }
  try {
    check.reference(field, body.replaceAll("$", ""), { range: true });
  } catch (error) {
    // Reported with the link as given: the part after the sheet name alone — often empty — named nothing the caller
    // wrote.
    if (error instanceof InvalidAddressError && error.address !== link) {
      // `Form.addCheckbox: Invalid address "B": …` — the call in front, the details after.
      const lead = `Invalid address "${error.address}": `;
      const at = error.message.indexOf(lead);
      const details = error.message.slice(at + lead.length);
      throw new InvalidAddressError(link, details, {
        cause: error,
        field: error.field,
        ...(at > 2 ? { where: error.message.slice(0, at - 2) } : {})
      });
    }
    throw error;
  }
}

/**
 * Add a legacy form checkbox over `range` — a cell (`"B2"`, sized to about two columns), a range (`"B2:D3"`), a
 * `{ tl, br }` pair or 0-based `{ startCol, startRow, endCol, endRow }` — and return it.
 *
 * @throws {InvalidAddressError} when `range` is not a reference.
 * @throws {ExcelArgumentError} when a position is not a whole number inside the grid, or an option is of the wrong type.
 */
export function addCheckbox(
  ws: WorksheetData,
  range: FormControlRange,
  options?: FormCheckboxOptions
): FormCheckboxData {
  checkWorksheet(ws);
  const check = argumentChecks("Form.addCheckbox");
  checkRange(check, range);
  if (check.record("options", options)) {
    checkLink(check, "options.link", options.link);
    check.string("options.text", options.text);
    check.boolean("options.checked", options.checked);
    check.boolean("options.noThreeD", options.noThreeD);
    check.boolean("options.print", options.print);
    check.oneOf("options.editAs", options.editAs, ["oneCell", "twoCell", "absolute"]);
  }
  return addFormCheckbox(ws, range, options);
}

/** Set the cell the checkbox writes TRUE/FALSE to (made absolute), or clear it with `undefined`. */
export function setLink(checkbox: FormCheckboxData, value: string | undefined): void {
  checkFormCheckbox(checkbox);
  checkLink(argumentChecks("Form.setLink"), "value", value);
  formCheckboxSetLink(checkbox, value);
}

/** Set the label shown next to the checkbox. */
export function setText(checkbox: FormCheckboxData, value: string): void {
  checkFormCheckbox(checkbox);
  argumentChecks("Form.setText").string("value", value);
  formCheckboxSetText(checkbox, value);
}

/** Check (`true`) or uncheck (`false`) the checkbox. */
export function setChecked(checkbox: FormCheckboxData, value: boolean): void {
  checkFormCheckbox(checkbox);
  const check = argumentChecks("Form.setChecked");
  check.required("value", value);
  check.boolean("value", value);
  formCheckboxSetChecked(checkbox, value);
}

/** Whether the checkbox is checked. */
export { formCheckboxChecked as isChecked };
/** The cell the checkbox writes TRUE/FALSE to, in absolute form (`$A$1`), or `undefined`. */
export { formCheckboxLink as getLink };
/** The label shown next to the checkbox. */
export { formCheckboxText as getText };

/** The model a VML getter reads: the handle's, or the model itself as the getters used to require. */
function modelOf(checkbox: FormCheckboxData | FormCheckboxModel): FormCheckboxModel {
  if (isFormCheckbox(checkbox)) {
    return checkbox.model;
  }
  // A model has its two anchors; anything else is neither, and is refused rather than read as one.
  const model = checkbox as Partial<FormCheckboxModel> | null | undefined;
  if (typeof model?.tl !== "object" || model.tl === null || typeof model.br !== "object") {
    checkFormCheckbox(checkbox);
  }
  return checkbox;
}

/** The VML `<x:Anchor>` text for a checkbox: its cell anchors and pixel offsets. */
export function getVmlAnchor(checkbox: FormCheckboxData): string;
/** {@link getVmlAnchor} for a checkbox model. */
export function getVmlAnchor(model: FormCheckboxModel): string;
export function getVmlAnchor(checkbox: FormCheckboxData | FormCheckboxModel): string {
  return formCheckboxVmlAnchor(modelOf(checkbox));
}

/** The VML positioning `style` for a checkbox. */
export function getVmlStyle(checkbox: FormCheckboxData): string;
/** {@link getVmlStyle} for a checkbox model. */
export function getVmlStyle(model: FormCheckboxModel): string;
export function getVmlStyle(checkbox: FormCheckboxData | FormCheckboxModel): string {
  return formCheckboxVmlStyle(modelOf(checkbox));
}

/** The VML `<x:Checked>` number for a checkbox: `0` unchecked, `1` checked, `2` mixed. */
export function getVmlCheckedValue(checkbox: FormCheckboxData): number;
/** {@link getVmlCheckedValue} for a checkbox model. */
export function getVmlCheckedValue(model: FormCheckboxModel): number;
export function getVmlCheckedValue(checkbox: FormCheckboxData | FormCheckboxModel): number {
  return formCheckboxVmlCheckedValue(modelOf(checkbox));
}

/** A form-checkbox handle. */
export type { FormCheckboxData as Handle } from "@excel/core/form-control";

// --- deprecated names ---

/**
 * @deprecated Use {@link addCheckbox}. This builds a checkbox for `worksheet` without adding it, so two calls take the
 * same shape id and neither is written. Removed in 1.0.
 */
export function createCheckbox(
  worksheet: WorksheetData,
  range: FormControlRange,
  options?: FormCheckboxOptions
): FormCheckboxData {
  checkWorksheet(worksheet);
  return formCheckboxCreate(worksheet, range, options);
}

/**
 * @deprecated Use {@link addCheckbox}. This builds a handle from a model without adding it to the sheet, so it is not
 * listed and not written, and nothing public attaches it; a checkbox's model is `handle.model`, which the VML getters
 * take as it is. Removed in 1.0.
 */
export function checkboxFromModel(
  worksheet: WorksheetData,
  model: FormCheckboxModel
): FormCheckboxData {
  checkWorksheet(worksheet);
  return formCheckboxFromModel(worksheet, model);
}
/** @deprecated Use {@link isChecked}. Removed in 1.0. */
export const checked = formCheckboxChecked;
/** @deprecated Use {@link getLink}. Removed in 1.0. */
export const link = formCheckboxLink;
/** @deprecated Use {@link getText}. Removed in 1.0. */
export const text = formCheckboxText;
/** @deprecated Use {@link getVmlAnchor}. Removed in 1.0. */
export const vmlAnchor = getVmlAnchor;
/** @deprecated Use {@link getVmlStyle}. Removed in 1.0. */
export const vmlStyle = getVmlStyle;
/** @deprecated Use {@link getVmlCheckedValue}. Removed in 1.0. */
export const vmlCheckedValue = getVmlCheckedValue;
/**
 * @deprecated Use {@link list}, which returns a copy. Removed in 1.0. This form
 * returns the sheet's live array, as it always did: before `remove` existed, splicing it was
 * how a checkbox was removed.
 */
export function listCheckboxes(ws: WorksheetData): FormCheckboxData[] {
  checkWorksheet(ws);
  return getFormCheckboxes(ws);
}
