/**
 * `Vba` namespace surface — VBA macro project inspection / management.
 *
 * `import { Vba } from "@cj-tech-master/excelts/word"` →
 *   `Vba.has(doc)`, `Vba.getInfo(doc)`, `Vba.add(doc, bin)`,
 *   `Vba.remove(doc)`, … — tree-shaken via `export * as Vba`.
 */
export {
  hasVbaProject as has,
  getVbaProjectInfo as getInfo,
  getVbaProjectData as getData,
  addVbaProject as add,
  removeVbaProject as remove,
  listVbaParts as listParts
} from "@word/advanced/vba-project";
