/**
 * `Form` namespace surface — legacy form controls.
 *
 * Worksheet-level:
 *   `Form.addCheckbox(ws, range, opts)`, `Form.listCheckboxes(ws)`.
 * Checkbox-handle operations:
 *   `Form.isCheckbox(v)`, `Form.checked(cb)`, `Form.setChecked(cb, true)`,
 *   `Form.link(cb)`, `Form.setText(cb, "Label")`, …
 */
export {
  addFormCheckbox as addCheckbox,
  getFormCheckboxes as listCheckboxes
} from "@excel/worksheet";

export {
  isFormCheckbox as isCheckbox,
  formCheckboxCreate as createCheckbox,
  formCheckboxFromModel as checkboxFromModel,
  formCheckboxChecked as checked,
  formCheckboxSetChecked as setChecked,
  formCheckboxLink as link,
  formCheckboxSetLink as setLink,
  formCheckboxText as text,
  formCheckboxSetText as setText,
  formCheckboxVmlAnchor as vmlAnchor,
  formCheckboxVmlStyle as vmlStyle,
  formCheckboxVmlCheckedValue as vmlCheckedValue
} from "@excel/form-control";

/** A form-checkbox handle. */
export type { FormCheckboxData as Handle } from "@excel/form-control";
