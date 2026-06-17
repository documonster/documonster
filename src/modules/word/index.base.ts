/**
 * DOCX Module - Base Exports (Platform Independent)
 *
 * Shared exports for both Node.js and browser environments.
 *
 * Each export is annotated with a `@stability` tag:
 *
 * - `@stability stable` — Public API. Backwards-compatible changes only;
 *   semver-major bumps for any breaking changes.
 * - `@stability experimental` — Public but evolving. The shape may change
 *   without a major version bump while we collect feedback. Consumers
 *   should pin to exact minor versions if they rely on these.
 * - `@stability preserve-only` — Round-trip preservation types. They exist
 *   so reading and writing a DOCX is loss-less; the underlying schema is
 *   not yet promised to be ergonomic. Avoid building features on top of
 *   them unless you also handle the underlying OOXML directly.
 *
 * Tree-shakeable subpaths: HTML/Markdown/Excel/PDF integrations live under
 * `excelts/word/{html,markdown,excel}` and `excelts/pdf` — they are not
 * re-exported here so bundlers can drop them when unused.
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
  EastAsianCombineBrackets,
  EastAsianLayoutSpec,
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
  OpaqueParagraphChild,
  OleObjectPart
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

// --- Stable API ---
/** @stability stable */
export * as Document from "./builder/document-handle";
/** @stability stable */
export type { DocumentHandle } from "./builder/document-handle";

// =============================================================================
// Domain namespaces — public value API (tree-shaken via `export * as`).
// Functions live in `surface/*.ts`; the shared data-model types stay exported
// flat from `./types` above. Mirrors the excel module's surface structure.
// =============================================================================
export * as Build from "./surface/build";
export * as Units from "./surface/units";
export * as Theme from "./surface/theme";
export * as Query from "./surface/query";
export * as Io from "./surface/io";
export * as Template from "./surface/template";
export * as Convert from "./surface/convert";
export * as Font from "./surface/font";
export * as Layout from "./surface/layout";
export * as Security from "./surface/security";
export * as Ole from "./surface/ole";
export * as Vba from "./surface/vba";
export * as Glossary from "./surface/glossary";
export * as Styles from "./surface/styles";
export * as Diff from "./surface/diff";
export * as Validation from "./surface/validation";
export * as Streaming from "./surface/streaming";
export * as RenderContext from "./surface/render-context";

/** @stability stable */
export type { SearchResult, DocumentHeading, DocumentSection } from "./query/search";

/** @stability stable */
export type { MergeOptions } from "./query/merge";

/** @stability stable */
export type { SplitOptions } from "./query/split";

/** @stability stable */
export type { RevisionEntry } from "./query/revisions";

/** @stability stable */
export type {
  StyleResolveContext,
  ResolvedParagraphStyle,
  ResolvedRunStyle,
  ResolvedNumberingLevel
} from "./query/style-resolve";

/** @stability stable */
export type { CompatibilityMode } from "./query/compat";

/** @stability stable */
export type { FormFieldEntry } from "./query/form-fields";

/** @stability stable */
export type { PatchContent, PatchOperation, PatchOptions, CompiledTemplate } from "./document-io";

// Incremental edit API (low-level efficient editing)
export type { IncrementalEdit, IncrementalEditOptions } from "./incremental-edit";

// --- Stable API ---

/** @stability stable */
export type { ReadDocxOptions } from "./reader/docx-reader";

// --- Experimental API ---
/** @stability experimental */
export type { DiffChangeType, DiffEntry, DiffSummary, DiffResult } from "./advanced/diff";

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

// --- Stable API ---
/** @stability stable */
export type { FlatOpcPart } from "./convert/flat-opc";

// --- Experimental API ---
/** @stability experimental */
export type { StreamingDocxOptions, StreamingProgressCallback } from "./writer/streaming-writer";

/** @stability experimental */
export type {
  ValidationSeverity,
  ValidationIssue,
  ValidationResult,
  ValidationOptions
} from "./advanced/validation";

// Font embedding
/** @stability stable */
export type { FontEmbedStyle, EmbedFontOptions, EmbedFontResult } from "./font/font-embed";

// --- Experimental API ---
/** @stability experimental */
export type {
  OleObject,
  OleObjectType,
  OleDisplayAs,
  OleExtractionResult,
  OleEmbeddingResult
} from "./advanced/ole-objects";

/** @stability experimental */
export type { BuildingBlockGallery, BuildingBlock, GlossaryDocument } from "./advanced/glossary";

// Document protection
/** @stability stable */
export type {
  ProtectionEditType,
  ProtectionHashAlgorithm,
  DocumentProtectionOptions,
  ProtectionState
} from "./security/document-protection";

// Style mapping DSL
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
export type { VbaProjectInfo } from "./advanced/vba-project";

/** @stability experimental */
export type { LayoutResult, LayoutOptions } from "./layout/layout";
/** @stability experimental */
export type { FullLayoutOptions, PageGeometryOverride } from "./layout/layout-full";

/** @stability experimental */
export type { FieldUpdateOptions } from "./advanced/field-engine";

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

export type { ScriptType, BiDiDirection, ShapedCluster, ShapingOptions } from "./font/text-shaping";

export type { HyphenationOptions, HyphenationPatterns } from "./font/hyphenation";

export type { FormatCriteria, FormatSearchResult } from "./query/format-search";

export type { RenderOptions } from "./layout/render-page";

export type { DataSource } from "./template/template-datasource";

export type { ChartBinding, ChartSeriesData, ChartTemplateData } from "./template/template-chart";

// NOTE: renderToMarkdown, renderToHtml, htmlToDocx, excelToDocx, markdownToDocx are NOT exported
// from the main entry — they live in dedicated subpaths to ensure tree-shaking:
//   import { renderToMarkdown, markdownToDocx } from "excelts/word/markdown"
//   import { renderToHtml }     from "excelts/word/html"
//   import { excelToDocx }      from "excelts/word/excel"
//   import { docxToPdf }        from "excelts/pdf"

// --- Stable API ---
export type { EncryptOptions } from "./security/encryption";

// --- Security Policy ---
/** @stability experimental */
export type { WordSecurityPolicy } from "./security/policy";

// --- Render Context (experimental) ---
/** @stability experimental */
export type { WordRenderContext, IdGenerators } from "./writer/render-context";

// --- Document Transformer (experimental) ---
/** @stability experimental */
export type { DocxTransformer, MapOptions } from "./core/mapper";

// --- Document Walker (experimental) ---
/** @stability experimental */
export type { DocxVisitor, WalkPath, WalkOptions, VisitAction } from "./core/walker";

// --- Layout Model (experimental) ---
/** @stability experimental */
export type {
  LayoutDocument,
  LayoutPage,
  PageGeometry,
  PageContent,
  LayoutParagraph,
  LayoutTable,
  LayoutTableCell,
  LayoutImage,
  LayoutFloat,
  LayoutFloatWrap,
  LayoutTextBox,
  LayoutShape,
  LayoutChart,
  LayoutSdt,
  LayoutMath,
  LayoutCheckBox,
  LayoutTableOfContents,
  LayoutAltChunk,
  LayoutOpaqueDrawing,
  LineBox,
  LineBoxItem,
  PositionedInlineImage,
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
export type { DocxToSemanticOptions } from "./convert/docx-to-semantic";

// --- Model helpers ---
/** @stability experimental */
export type { AssetRef, MediaAsset } from "./types";
