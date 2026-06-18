/**
 * `Glossary` namespace surface — building blocks / glossary document
 * (AutoText, Quick Parts).
 *
 * `import { Glossary } from "documonster/word"` →
 *   `Glossary.createBlock(...)`, `Glossary.findBlock(...)`,
 *   `Glossary.autoTextEntries(g)`, … — tree-shaken via `export * as Glossary`.
 */
export {
  createBuildingBlock as createBlock,
  createGlossaryDocument as createDocument,
  findBuildingBlock as findBlock,
  listBuildingBlocks as listBlocks,
  getAutoTextEntries as autoTextEntries,
  getQuickParts as quickParts
} from "@word/advanced/glossary";
