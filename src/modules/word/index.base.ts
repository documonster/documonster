/**
 * DOCX Module - Base Exports (Platform Independent)
 *
 * Shared exports for both Node.js and browser environments.
 */

// --- Stable API ---
/** @stability stable */
export type {
  // Units
  HalfPoint,
  Twips,
  Emu,
  EighthPoint,
  HexColor,

  // Shading & Border enums
  ShadingType,
  BorderStyle,

  // Page & Section
  PageOrientation,
  PageSize,
  PageMargins,
  ColumnDef,
  SectionColumns,
  SectionBreakType,
  PageNumberFormat,
  HeaderFooterType,
  HeaderFooterRef,
  PageVerticalAlign,
  DocumentGridType,
  PageTextDirection,
  Border,
  ArtBorderType,
  PageBorders,
  SectionProperties,

  // Run
  UnderlineStyle,
  VerticalAlign,
  HighlightColor,
  FontSpec,
  Shading,
  UnderlineSpec,
  TextEffect,
  EmphasisMarkType,
  ColorSpec,
  RunProperties,

  // Run Content
  RunContent,
  TextContent,
  BreakContent,
  TabContent,
  PositionalTabContent,
  RubyContent,
  RubyAlign,
  RubyProperties,
  SymbolContent,
  FootnoteRefContent,
  EndnoteRefContent,
  FieldContent,
  InlineImageContent,
  CarriageReturnContent,
  NoBreakHyphenContent,
  SoftHyphenContent,
  LastRenderedPageBreakContent,
  AnnotationReferenceContent,
  PageNumberType,
  DateFieldContent,
  Run,

  // Track Changes
  RevisionInfo,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  MoveRangeMarker,
  CustomXmlTrackingMarker,
  ParagraphPropertyChange,
  RunPropertyChange,
  SectionPropertyChange,
  TableRowRevision,

  // Comments
  CommentDef,
  CommentRangeStart,
  CommentRangeEnd,
  CommentReference,

  // Paragraph
  Alignment,
  LineSpacingRule,
  LineSpacing,
  Indentation,
  TabStopType,
  TabStopLeader,
  TabStop,
  PositionalTabAlignment,
  PositionalTabRelativeTo,
  PositionalTabLeader,
  ParagraphBorders,
  NumberingRef,
  FrameAnchorType,
  DropCapType,
  ParagraphFrame,
  ParagraphProperties,
  Paragraph,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  ParagraphChild,

  // Table
  TableWidth,
  TableBorders,
  TableCellMargins,
  TableLayout,
  TableLook,
  TableFloat,
  TableProperties,
  RowHeightRule,
  TableRowProperties,
  VerticalCellAlign,
  TextDirection,
  VerticalMerge,
  TableCellProperties,
  TableCell,
  TableRow,
  Table,

  // Floating Image
  HorizontalPositionRelative,
  VerticalPositionRelative,
  WrapStyle,
  WrapTextSide,
  WrapMargins,
  FloatingImage,

  // Drawing shapes
  DrawingShape,
  PresetShapeType,

  // Charts
  Chart,
  ChartType,
  ChartSeries,
  ChartLegendPosition,
  ChartAxis,
  ChartContent,
  ChartExContent,
  ChartExData,
  ChartExSeriesData,
  ChartTrendline,
  ChartTrendlineType,
  ChartErrorBars,
  ChartErrorBarDirection,
  ChartErrorBarType,
  ChartDataLabels,
  ChartDataLabelPosition,

  // Alt chunks (embedded HTML/RTF content)
  AltChunk,

  // Web settings & people (collaboration metadata)
  WebSettings,
  PersonInfo,

  // Notes type enum
  NoteType,
  FootnoteProperties,
  EndnoteProperties,
  NoteNumberFormat,
  NoteNumberRestart,

  // Table revision types
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  CellMergeRevision,

  // Document type
  DocxDocumentType,

  // TOC
  TableOfContents,

  // Math
  MathRun,
  MathFraction,
  MathSuperScript,
  MathSubScript,
  MathSubSuperScript,
  MathPreSubSuperScript,
  MathPhantom,
  MathGroupChar,
  MathBorderBox,
  MathRadical,
  MathDelimiter,
  MathNary,
  MathFunction,
  MathLimit,
  MathMatrix,
  MathAccent,
  MathBar,
  MathBox,
  MathEquationArray,
  MathContent,
  MathBlock,

  // TextBox
  TextBox,

  // CheckBox
  CheckBox,

  // SDT
  SdtListItem,
  SdtDateProperties,
  SdtProperties,
  SdtDataBinding,
  SdtCheckboxProperties,
  SdtRepeatingSectionProperties,
  StructuredDocumentTag,
  CustomXmlPart,

  // Form Fields
  FormFieldType,
  FormField,
  TextFormField,
  CheckBoxFormField,
  DropDownFormField,

  // Watermark
  Watermark,
  TextWatermark,
  ImageWatermark,

  // Body
  BodyContent,

  // Styles
  StyleType,
  TableStyleConditionType,
  TableStyleConditionalFormat,
  StyleDef,
  DocDefaults,

  // Numbering
  NumberFormat,
  LevelJustification,
  LevelSuffix,
  NumberingLevel,
  MultiLevelType,
  AbstractNumbering,
  LevelOverride,
  NumberingInstance,
  NumPicBullet,

  // Headers/Footers
  HeaderFooterContent,
  HeaderDef,
  FooterDef,

  // Footnotes/Endnotes
  FootnoteDef,
  EndnoteDef,

  // Images
  ImageMediaType,
  ImageDef,

  // Font Table
  FontFamily,
  FontPitch,
  FontDef,
  EmbeddedFont,

  // Settings
  ProtectionType,
  HyphenationSettings,
  DocumentSettings,
  CompatSetting,
  CompatFlag,

  // Document Background
  DocumentBackground,
  DocumentTheme,
  ThemeColorScheme,
  ThemeFontScheme,
  ThemeFont,
  ThemeFormatScheme,
  ThemeColorName,

  // Custom Properties
  CustomPropertyValue,
  CustomProperty,

  // Core/App Properties
  CoreProperties,
  AppProperties,

  // Document
  DocxDocument,
  DocxOptions
} from "./types";

// --- Preserve-only API (round-trip preservation, no full editing support) ---
/** @stability preserve-only */
export type {
  OpaquePart,
  OpaqueRelationship,
  OpaqueDrawing,
  OpaqueRunContent,
  OpaqueParagraphChild
} from "./types";

// --- Stable API ---
/** @stability stable */
export {
  DocxError,
  DocxParseError,
  DocxWriteError,
  DocxMissingPartError,
  DocxInvalidStructureError,
  DocxUnsupportedFeatureError,
  DocxEncryptedError,
  DocxDecryptionError,
  DocxLimitExceededError,
  isDocxError
} from "./errors";

// --- Stable API ---
/** @stability stable */
export {
  inchesToTwips,
  twipsToInches,
  cmToTwips,
  twipsToCm,
  ptToTwips,
  twipsToPt,
  mmToTwips,
  inchesToEmu,
  emuToInches,
  cmToEmu,
  emuToCm,
  ptToEmu,
  pxToEmu,
  emuToPx,
  ptToHalfPoint,
  halfPointToPt,
  ptToEighthPoint,
  eighthPointToPt,
  lineMultiplierToSpacing,
  spacingToLineMultiplier,
  percentToTablePct,
  tablePctToPercent
} from "./units";

// --- Stable API ---
/** @stability stable */
export {
  Document,
  // Run helpers
  text,
  bold,
  italic,
  underline,
  strikethrough,
  pageBreak,
  lineBreak,
  columnBreak,
  tab,
  positionalTab,
  ruby,
  carriageReturn,
  noBreakHyphen,
  softHyphen,
  field,
  pageNumberField,
  totalPagesField,
  sectionPagesField,
  sectionField,
  dateField,
  sequenceField,
  timeField,
  authorField,
  titleField,
  subjectField,
  keywordsField,
  fileNameField,
  fileSizeField,
  styleRefField,
  refField,
  pageRefField,
  noteRefField,
  hyperlinkField,
  quoteField,
  tocField,
  tcField,
  indexEntryField,
  indexField,
  ifField,
  includeTextField,
  includePictureField,
  formTextField,
  formCheckboxField,
  formDropdownField,
  // Paragraph helpers
  paragraph,
  textParagraph,
  heading,
  hyperlink,
  bookmarkStart,
  bookmarkEnd,
  // Comment helpers
  commentRangeStart,
  commentRangeEnd,
  commentReference,
  // Track changes helpers
  insertedRun,
  deletedRun,
  movedFromRun,
  movedToRun,
  moveFromRangeStart,
  moveFromRangeEnd,
  moveToRangeStart,
  moveToRangeEnd,
  // CheckBox helper
  checkBox,
  // Math helpers
  mathBlock,
  mathRun,
  mathFraction,
  mathSqrt,
  mathRoot,
  mathSum,
  mathIntegral,
  mathProduct,
  // Additional math helpers
  mathSuperScript,
  mathSubScript,
  mathSubSuperScript,
  mathPreSubSuperScript,
  mathPhantom,
  mathGroupChar,
  mathBorderBox,
  mathDelimiter,
  mathNary,
  mathFunction,
  mathLimit,
  mathMatrix,
  mathAccent,
  mathBar,
  mathBox,
  mathEquationArray,
  // Symbol helper
  symbol,
  // Floating image helper
  floatingImage,
  // Drawing shape helper
  drawingShape,
  // Chart helper
  chart,
  // SDT helper
  structuredDocumentTag,
  // Table helpers
  border,
  gridBorders,
  cell,
  row,
  table,
  simpleTable,
  searchText,
  replaceText,
  resolveThemeColor,
  mailMerge,
  // Document merge
  mergeDocuments,
  // Style resolution
  resolveStyle,
  resolveRunStyle,
  resolveNumberingLevel,
  resolveTableStyle,
  // Compatibility mode
  getCompatibilityMode,
  setCompatibilityMode,
  // Track changes accept/reject
  acceptAllRevisions,
  rejectAllRevisions,
  listRevisions,
  acceptRevision,
  rejectRevision,
  // Document split
  splitDocument,
  // OpenDoPE data binding
  resolveDataBindings,
  // Form field operations
  extractFormFields,
  fillFormFields,
  // Query API
  paragraphCount,
  countWords,
  getHeadings,
  findBookmark,
  findComment,
  listImages,
  listTables,
  listHyperlinks,
  listSections,
  tableCount,
  extractText
} from "./document";
/** @stability stable */
export type {
  DocumentHandle,
  SearchResult,
  DocumentHeading,
  DocumentSection,
  MergeOptions,
  ResolvedParagraphStyle,
  ResolvedRunStyle,
  ResolvedNumberingLevel,
  StyleResolveContext,
  CompatibilityMode,
  FormFieldEntry,
  RevisionEntry,
  SplitOptions
} from "./document";
/** @stability stable */
export type { PatchContent, PatchOperation, PatchOptions, CompiledTemplate } from "./document-io";

// Incremental edit API (low-level efficient editing)
export type { IncrementalEdit, IncrementalEditOptions } from "./incremental-edit";
export { editDocxIncremental, listDocxParts, readDocxPart } from "./incremental-edit";

// --- Stable API ---
/** @stability stable */
export { packageDocx } from "./writer/docx-packager";

/** @stability stable */
export { readDocx } from "./reader/docx-reader";
/** @stability stable */
export type { ReadDocxOptions } from "./reader/docx-reader";

/** @stability stable */
export {
  toBuffer,
  toBase64,
  patchDocument,
  compileTemplate,
  patchTemplate,
  fillTemplateFromBuffer,
  toFlatOpcFromDoc
} from "./document-io";

// --- Experimental API ---
/** @stability experimental */
export { diffDocuments } from "./advanced/diff";
/** @stability experimental */
export type { DiffChangeType, DiffEntry, DiffSummary, DiffResult } from "./advanced/diff";

/** @stability experimental */
export {
  fillTemplate,
  fillTemplateEnhanced,
  listTemplateTags,
  isTemplateChart,
  TemplateError
} from "./template/template-engine";
/** @stability experimental */
export type {
  TemplateOptions,
  TemplateImage,
  TemplateRichText,
  TemplateSubDocument,
  TemplateChart,
  TemplateHtmlChunk,
  TemplateTag
} from "./template/template-engine";

// Sub-namespaces for grouped API ergonomics (no name conflicts with types)
/** @stability stable */
export { Field, Drawing, TrackChanges, Sdt, Query, Comment } from "./namespaces";
/** @stability stable */
export { Math as MathML } from "./namespaces";

// --- Stable API ---
/** @stability stable */
export { parseFlatOpc, isFlatOpc, toFlatOpc } from "./convert/flat-opc";
/** @stability stable */
export type { FlatOpcPart } from "./convert/flat-opc";

// --- Experimental API ---
/** @stability experimental */
export { StreamingDocxWriter, createDocxStream } from "./writer/streaming-writer";
/** @stability experimental */
export type { StreamingDocxOptions, StreamingProgressCallback } from "./writer/streaming-writer";

/** @stability experimental */
export { validateDocument } from "./advanced/validation";
/** @stability experimental */
export type {
  ValidationSeverity,
  ValidationIssue,
  ValidationResult,
  ValidationOptions
} from "./advanced/validation";

// Font embedding
/** @stability stable */
export { embedFont, embedFontFamily, addEmbeddedFonts, subsetFont } from "./font/font-embed";
/** @stability stable */
export type { FontEmbedStyle, EmbedFontOptions, EmbedFontResult } from "./font/font-embed";

// --- Experimental API ---
/** @stability experimental */
export {
  extractOleObjects,
  hasOleObjects,
  getOleObjectData,
  createOleEmbedding
} from "./advanced/ole-objects";
/** @stability experimental */
export type {
  OleObject,
  OleObjectType,
  OleDisplayAs,
  OleExtractionResult
} from "./advanced/ole-objects";

/** @stability experimental */
export {
  createBuildingBlock,
  createGlossaryDocument,
  findBuildingBlock,
  listBuildingBlocks,
  getAutoTextEntries,
  getQuickParts
} from "./advanced/glossary";
/** @stability experimental */
export type { BuildingBlockGallery, BuildingBlock, GlossaryDocument } from "./advanced/glossary";

// Document protection
/** @stability stable */
export {
  protectDocument,
  unprotectDocument,
  isDocumentProtected,
  getProtectionState,
  verifyProtectionPassword
} from "./security/document-protection";
/** @stability stable */
export type {
  ProtectionEditType,
  ProtectionHashAlgorithm,
  DocumentProtectionOptions,
  ProtectionState
} from "./security/document-protection";

// Style mapping DSL
/** @stability stable */
export {
  parseStyleMap,
  createStyleMap,
  mergeStyleMaps,
  matchStyleMap,
  DEFAULT_STYLE_MAP
} from "./advanced/style-map";
/** @stability stable */
export type {
  MappingSourceType,
  MappingCondition,
  MappingTarget,
  StyleMappingRule,
  StyleMap,
  StyleMapOptions
} from "./advanced/style-map";

// --- Experimental API ---
/** @stability experimental */
export {
  hasVbaProject,
  getVbaProjectInfo,
  getVbaProjectData,
  addVbaProject,
  removeVbaProject,
  listVbaParts
} from "./advanced/vba-project";
/** @stability experimental */
export type { VbaProjectInfo } from "./advanced/vba-project";

/** @stability experimental */
export { layoutDocument } from "./layout/layout";
/** @stability experimental */
export type { LayoutResult, LayoutOptions } from "./layout/layout";
/** @stability experimental */
export { layoutDocumentFull } from "./layout/layout-full";
/** @stability experimental */
export type { FullLayoutOptions } from "./layout/layout-full";

/** @stability experimental */
export { updateFields, updateTableOfContents } from "./advanced/field-engine";
/** @stability experimental */
export type { FieldUpdateOptions } from "./advanced/field-engine";

/** @stability experimental */
export { ommlToMathML, mathMLToOmml } from "./advanced/math-convert";

/** @stability experimental */
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
} from "./advanced/drawing-shapes";
/** @stability experimental */
export type {
  StandardShapeType,
  SolidFill,
  GradientStop,
  GradientFill,
  PatternFill,
  NoFill,
  ShapeFill,
  LineDash,
  LineEndType,
  LineEndSize,
  ShapeOutline,
  ShadowEffect,
  GlowEffect,
  ReflectionEffect,
  Effect3D,
  ShapeEffects,
  TextVerticalAnchor,
  TextWrap,
  ShapeTextBody,
  CreateShapeOptions
} from "./advanced/drawing-shapes";

/** @stability experimental — Text shaping for complex scripts */
export { shapeText, detectScript, detectDirection } from "./font/text-shaping";
export type { ScriptType, BiDiDirection, ShapedCluster, ShapingOptions } from "./font/text-shaping";

/** @stability experimental — Hyphenation engine */
export {
  createHyphenator,
  hyphenateWord,
  hyphenateText,
  ENGLISH_US_PATTERNS
} from "./font/hyphenation";
export type { HyphenationOptions, HyphenationPatterns } from "./font/hyphenation";

/** @stability experimental — Format-based search */
export { searchByFormat, countByFormat, getUsedFormats } from "./query/format-search";
export type { FormatCriteria, FormatSearchResult } from "./query/format-search";

/** @stability experimental — ODT (OpenDocument Text) read/write */
export { readOdt, writeOdt } from "./convert/odt/odt";

/** @stability experimental — Page rendering to SVG */
export { renderPageToSvg, renderDocumentToSvg, renderPageFromLayout } from "./layout/render-page";
export type { RenderOptions } from "./layout/render-page";

/** @stability experimental — Template data source abstraction */
export {
  JsonDataSource,
  XmlDataSource,
  CsvDataSource,
  CompositeDataSource,
  fillTemplateFromSource
} from "./template/template-datasource";
export type { DataSource } from "./template/template-datasource";

/** @stability experimental — Chart data binding for templates */
export { bindChartData } from "./template/template-chart";
export type { ChartBinding, ChartSeriesData, ChartTemplateData } from "./template/template-chart";

// NOTE: renderToMarkdown, renderToHtml, htmlToDocx, excelToDocx, markdownToDocx are NOT exported
// from the main entry — they live in dedicated subpaths to ensure tree-shaking:
//   import { renderToMarkdown, markdownToDocx } from "excelts/word/markdown"
//   import { renderToHtml }     from "excelts/word/html"
//   import { excelToDocx }      from "excelts/word/excel"
//   import { docxToPdf }        from "excelts/pdf"

// --- Stable API ---
/** @stability stable — DOCX encryption */
export { encryptDocx } from "./security/encryption";
export type { EncryptOptions } from "./security/encryption";

// --- OPC Package Model (experimental) ---
/** @stability experimental */
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
} from "./core/opc-package";
/** @stability experimental */
export { planToPackage, normalizePartName, resolveTarget } from "./core/opc-package";

// --- Security Policy ---
/** @stability experimental */
export type { WordSecurityPolicy } from "./security/policy";
/** @stability experimental */
export {
  DEFAULT_SECURITY_POLICY,
  STRICT_SECURITY_POLICY,
  resolveSecurityPolicy
} from "./security/policy";

// --- Render Context (experimental) ---
/** @stability experimental */
export type { WordRenderContext, IdGenerators } from "./writer/render-context";
/** @stability experimental */
export { createRenderContext, createIdGenerators } from "./writer/render-context";

// --- Document Transformer (experimental) ---
/** @stability experimental */
export type { DocxTransformer, MapOptions } from "./core/mapper";

export { mapDocument } from "./core/mapper";

// --- Layout Model (experimental) ---
/** @stability experimental */
export type {
  LayoutDocument,
  LayoutPage,
  PageGeometry,
  LayoutParagraph,
  LayoutTable,
  LayoutTableCell,
  LayoutImage,
  LayoutFloat,
  LineBox,
  PositionedRun
} from "./layout/layout-model";

// --- Conversion IR (experimental) ---
/** @stability experimental */
export type {
  SemanticDocument,
  SemanticBlock,
  SemanticInline,
  ConversionContext,
  ConversionWarning,
  ConversionAsset,
  ResolvedFormatting
} from "./convert/conversion-ir";
/** @stability experimental */
export { createConversionContext } from "./convert/conversion-ir";
/** @stability experimental */
export { docxToSemantic } from "./convert/docx-to-semantic";
/** @stability experimental */
export type { DocxToSemanticOptions } from "./convert/docx-to-semantic";

// --- Model helpers ---
/** @stability experimental */
export type { AssetRef, MediaAsset } from "./types";
