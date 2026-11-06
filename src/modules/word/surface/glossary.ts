import { getAutoTextEntries, getQuickParts } from "@word/advanced/glossary";
/**
 * `Glossary` namespace surface — building blocks / glossary document
 * (AutoText, Quick Parts).
 *
 * `import { Glossary } from "documonster/word"` →
 *   `Glossary.createBlock(...)`, `Glossary.findBlock(...)`,
 *   `Glossary.listAutoTextEntries(g)`, … — tree-shaken via
 *   `export * as Glossary`.
 */

export {
  createBuildingBlock as createBlock,
  createGlossaryDocument as createDocument,
  findBuildingBlock as findBlock,
  listBuildingBlocks as listBlocks,
  getAutoTextEntries as listAutoTextEntries,
  getQuickParts as listQuickParts
} from "@word/advanced/glossary";

/** @deprecated Use {@link listAutoTextEntries}. Removed in 1.0. */
export const autoTextEntries = getAutoTextEntries;
/** @deprecated Use {@link listQuickParts}. Removed in 1.0. */
export const quickParts = getQuickParts;
