/**
 * `DefinedNames` namespace surface — workbook/worksheet defined names.
 *
 * `import { DefinedNames } from "documonster/excel"` →
 *   `DefinedNames.create()`, `DefinedNames.add(dn, "Sheet1!A1", "MyName")`,
 *   `DefinedNames.addFormula(dn, "Tax", "0.2")`, `DefinedNames.remove(...)`.
 *
 * Exposes the high-level operations; lower-level `*Ex` / matrix helpers stay
 * internal (import from `@excel/defined-names` directly if needed in tests).
 */
export {
  createDefinedNames as create,
  definedNamesAdd as add,
  definedNamesAddFormula as addFormula,
  definedNamesRemove as remove,
  definedNamesGetNames as getNames,
  definedNamesGetAllEntries as getAllEntries,
  definedNamesModel as model,
  definedNamesSetModel as setModel
} from "@excel/defined-names";

/** A defined-names handle. */
export type { DefinedNamesData as Handle } from "@excel/defined-names";
