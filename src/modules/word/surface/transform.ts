/**
 * `Transform` namespace surface — whole-document transforms: merging,
 * splitting, field / TOC refresh, node mapping, form-field filling and data
 * binding.
 *
 * Every `Transform` function leaves the document it is given unchanged and
 * returns a new one. The result may share unchanged nodes with the input (or
 * be the input itself when there is nothing to change), so before editing a
 * result in place with `Document.*` while you still use the input, take
 * `Document.clone(result)`.
 *
 * `import { Transform } from "documonster/word"` →
 *   `Transform.merge([a, b])`, `Transform.split(doc)`,
 *   `Transform.updateFields(doc)`, `Transform.map(doc, transformer)`,
 *   `Transform.fillFormFields(doc, values)` — tree-shaken via
 *   `export * as Transform`.
 */
export { mergeDocuments as merge } from "@word/query/merge";
export { splitDocument as split } from "@word/query/split";
export { updateFields, updateTableOfContents } from "@word/advanced/field-engine";
export { mapDocument as map } from "@word/core/mapper";
export { fillFormFields } from "@word/query/form-fields";
export { resolveDataBindings } from "@word/query/data-binding";
