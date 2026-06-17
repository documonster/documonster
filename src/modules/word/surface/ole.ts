/**
 * `Ole` namespace surface — OLE embedded object extraction / embedding.
 *
 * `import { Ole } from "documonster/word"` →
 *   `Ole.extract(doc)`, `Ole.has(doc)`, `Ole.getData(doc, rId)`,
 *   `Ole.add(...)`, … — tree-shaken via `export * as Ole`.
 */
export {
  extractOleObjects as extract,
  hasOleObjects as has,
  getOleObjectData as getData,
  createOleEmbedding as createEmbedding,
  addOleObject as add
} from "../advanced/ole-objects";
