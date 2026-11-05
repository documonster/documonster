import { checkFormCheckbox } from "@excel/core/feature-handles";
import type { FormCheckboxData, FormCheckboxModel } from "@excel/core/form-control";
import {
  formCheckboxVmlAnchor,
  formCheckboxVmlCheckedValue,
  formCheckboxVmlStyle,
  isFormCheckbox
} from "@excel/core/form-control";
/**
 * `Form.Vml` — the VML text a legacy checkbox is written as. Supported API, but for a caller producing or checking the
 * VML part itself; building a sheet needs only `Form`.
 *
 * `import { Form } from "documonster/excel"` → `Form.Vml.getAnchor(cb)`, `Form.Vml.getStyle(cb)`,
 * `Form.Vml.getCheckedValue(cb)`.
 */

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
export function getAnchor(checkbox: FormCheckboxData): string;
/** {@link getAnchor} for a checkbox model. */
export function getAnchor(model: FormCheckboxModel): string;
export function getAnchor(checkbox: FormCheckboxData | FormCheckboxModel): string {
  return formCheckboxVmlAnchor(modelOf(checkbox));
}

/** The VML positioning `style` for a checkbox. */
export function getStyle(checkbox: FormCheckboxData): string;
/** {@link getStyle} for a checkbox model. */
export function getStyle(model: FormCheckboxModel): string;
export function getStyle(checkbox: FormCheckboxData | FormCheckboxModel): string {
  return formCheckboxVmlStyle(modelOf(checkbox));
}

/** The VML `<x:Checked>` number for a checkbox: `0` unchecked, `1` checked, `2` mixed. */
export function getCheckedValue(checkbox: FormCheckboxData): number;
/** {@link getCheckedValue} for a checkbox model. */
export function getCheckedValue(model: FormCheckboxModel): number;
export function getCheckedValue(checkbox: FormCheckboxData | FormCheckboxModel): number {
  return formCheckboxVmlCheckedValue(modelOf(checkbox));
}
