/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/advanced"` */

export { diffDocuments } from "./diff";
export type { DiffResult, DiffEntry, DiffSummary, DiffChangeType } from "./diff";
export { updateFields, updateTableOfContents } from "./field-engine";
export type { FieldUpdateOptions } from "./field-engine";
export { validateDocument } from "./validation";
export type {
  ValidationResult,
  ValidationIssue,
  ValidationSeverity,
  ValidationOptions
} from "./validation";
export {
  hasVbaProject,
  getVbaProjectInfo,
  getVbaProjectData,
  addVbaProject,
  removeVbaProject,
  listVbaParts
} from "./vba-project";
export type { VbaProjectInfo } from "./vba-project";
export {
  extractOleObjects,
  hasOleObjects,
  getOleObjectData,
  createOleEmbedding
} from "./ole-objects";
export type { OleObject, OleExtractionResult } from "./ole-objects";
export {
  createBuildingBlock,
  createGlossaryDocument,
  findBuildingBlock,
  listBuildingBlocks,
  getAutoTextEntries,
  getQuickParts
} from "./glossary";
export type { BuildingBlock, GlossaryDocument, BuildingBlockGallery } from "./glossary";
export {
  createShape,
  createRect,
  createRoundRect,
  createEllipse,
  createLine,
  createArrow,
  createFlowchartShape,
  createCallout,
  createStar
} from "./drawing-shapes";
export type { CreateShapeOptions, StandardShapeType } from "./drawing-shapes";
export {
  parseStyleMap,
  createStyleMap,
  mergeStyleMaps,
  matchStyleMap,
  DEFAULT_STYLE_MAP
} from "./style-map";
export type { StyleMap, StyleMappingRule, MappingTarget } from "./style-map";
export { ommlToMathML, mathMLToOmml } from "./math-convert";
export { searchByFormat, countByFormat, getUsedFormats } from "../query/format-search";
export type { FormatCriteria, FormatSearchResult } from "../query/format-search";
