/**
 * DOCX Module - Core (internal kernel)
 *
 * Re-exports shared utilities, type guards, text extraction, and the unified walker.
 * This barrel is for internal module use only — not part of the public API.
 */

export {
  bytesToBase64,
  base64ToBytes,
  generateUuid,
  generateGuid,
  stringToUtf16LE,
  utf16LEToString,
  buildAttrs,
  escapeXml,
  randomBytes
} from "./internal-utils";
export type { Mutable, DeepMutable } from "./internal-utils";

export {
  isParagraph,
  isTable,
  isSdt,
  isMathBlock,
  isTextBox,
  isFloatingImage,
  isRun,
  isHyperlink,
  isInsertedRun,
  isDeletedRun,
  isMovedFromRun,
  isMovedToRun,
  extractRunText,
  extractParagraphText,
  extractChildText,
  extractMathText,
  extractBodyText,
  collectBlockText
} from "./text-utils";

export type { VisitAction, WalkPath, WalkOptions, DocxVisitor } from "./walker";
export { walkDocument, walkBlocks, collectParagraphs, collectRuns, collectTables } from "./walker";

export type { DocxTransformer, MapOptions } from "./mapper";
export { mapDocument } from "./mapper";

export type {
  PartName,
  TargetMode,
  OpcRelationship,
  OpcRelationshipSet,
  OpcPart,
  ContentTypeEntry,
  OpcPackage,
  OpcWriteOptions,
  WordPackagePlan
} from "./opc-package";
export {
  planToPackage,
  normalizePartName,
  resolveTarget,
  getPartRelsPath,
  getFileName,
  getFileExt
} from "./opc-package";
