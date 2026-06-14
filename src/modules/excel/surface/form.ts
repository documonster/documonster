/**
 * `Form` namespace surface — legacy form controls.
 * `import { Form } from "documonster/excel"` → `Form.addCheckbox(ws, range, opts)`.
 */
export {
  addFormCheckbox as addCheckbox,
  getFormCheckboxes as listCheckboxes
} from "@excel/worksheet";
