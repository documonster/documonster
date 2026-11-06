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
 * `documonster/word/{html,markdown,excel}` and `documonster/pdf` — they are not
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
  FieldRangeMarker,
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
  RangeMarker,
  PermStart,
  PermEnd,
  BlockRangeMarkers,
  ParagraphChildRangeMarkers,
  CustomXmlTrackingMarker,
  BidiStart,
  BidiEnd,
  SmartTagStart,
  SmartTagEnd,
  SmartTagAttribute,
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
  HyperlinkNestedChild,
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
  NoteSeparatorType,
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
  SdtDocPart,
  BlockContentControl,
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
  PreservedXml,
  PreservedLink,
  CompatSetting,
  CompatFlag,

  // Document Background
  DocumentBackground,
  DocumentTheme,
  ThemeColorScheme,
  ThemeFontScheme,
  ThemeFont,
  ThemeSystemColor,
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
  DocxOptions,
  DocxDiscardUnrecognizedValuesOptions,
  DocxUnrecognizedValue
} from "@word/types";

// --- Preserve-only API (round-trip preservation, no full editing support) ---
/** @stability preserve-only */
export type {
  OpaquePart,
  OpaqueRelationship,
  OpaqueDrawing,
  OpaqueRunContent,
  OpaqueParagraphChild,
  OleObjectPart
} from "@word/types";

// --- Stable API ---
/** @stability stable */
export {
  DocxError,
  DocxParseError,
  DocxWriteError,
  DocxFileError,
  DocxDuplicateIdError,
  DocxInvalidImageError,
  DocxArgumentError,
  DocxInvalidDrawingError,
  DocxMissingPartError,
  DocxInvalidStructureError,
  DocxUnsupportedFeatureError,
  DocxEncryptedError,
  DocxDecryptionError,
  DocxLimitExceededError,
  DocxRawXmlPolicyError,
  DocxTemplateError,
  // Deprecated pre-prefix name; removed in 1.0.
  TemplateError,
  isDocxError
} from "@word/errors";
export type { DocxIdCollection } from "@word/errors";

// --- Stable API ---
/** @stability stable */
export * as Document from "@word/surface/document";
/** @stability stable */
export { PageSizes } from "@word/constants";
export type { DocxPageSizeName } from "@word/constants";
export type {
  DocxAddImageOptions,
  DocxAddFloatingImageOptions
} from "@word/builder/document-handle";
import type { DocxDocument } from "@word/types";
/** @deprecated Use {@link DocxDocument}. Removed in 1.0. */
export type DocumentHandle = DocxDocument;

// =============================================================================
// Domain namespaces — public value API (tree-shaken via `export * as`).
// Functions live in `surface/*.ts`; the shared data-model types stay exported
// flat from `./types` above. Mirrors the excel module's surface structure.
// =============================================================================
export * as Build from "@word/surface/build";
export * as Units from "@word/surface/units";
export * as Theme from "@word/surface/theme";
export * as Query from "@word/surface/query";
/** @deprecated Use `Document` (`Document.read`, `Document.toBuffer`, …). Removed in 1.0. */
export * as Io from "@word/surface/io";
export * as Package from "@word/surface/package";
export * as Transform from "@word/surface/transform";
export * as Template from "@word/surface/template";
export * as Convert from "@word/surface/convert";
export * as Font from "@word/surface/font";
export * as Layout from "@word/surface/layout";
export * as Security from "@word/surface/security";
export * as Ole from "@word/surface/ole";
export * as Vba from "@word/surface/vba";
export * as Glossary from "@word/surface/glossary";
export * as Styles from "@word/surface/styles";
export * as Diff from "@word/surface/diff";
export * as Validation from "@word/surface/validation";
export * as Stream from "@word/surface/streaming";
/** @deprecated Use {@link Stream}. Removed in 1.0. */
export * as Streaming from "@word/surface/streaming";
export * as RenderContext from "@word/surface/render-context";

/** @stability stable */
import type {
  SearchResult as DocxSearchResult,
  DocumentHeading as DocxHeading,
  DocumentSection as DocxSectionInfo,
  BookmarkLocation as DocxBookmarkLocation,
  ListTablesOptions as DocxListTablesOptions
} from "@word/query/search";
export type {
  DocxSearchResult,
  DocxHeading,
  DocxSectionInfo,
  DocxBookmarkLocation,
  DocxListTablesOptions
};

/** @stability stable */
import type { MergeOptions as DocxMergeOptions } from "@word/query/merge";
export type { DocxMergeOptions };

/** @stability stable */
import type { SplitOptions as DocxSplitOptions } from "@word/query/split";
export type { DocxSplitOptions };

/** @stability stable */
import type { RevisionEntry as DocxRevisionEntry } from "@word/query/revisions";
export type { DocxRevisionEntry };

/** @stability stable */
import type {
  StyleResolveContext as DocxStyleResolveContext,
  ResolvedParagraphStyle as DocxResolvedParagraphStyle,
  ResolvedRunStyle as DocxResolvedRunStyle,
  ResolvedTableStyle as DocxResolvedTableStyle,
  ResolvedNumberingLevel as DocxResolvedNumberingLevel,
  StyleIndex as DocxStyleIndex
} from "@word/query/style-resolve";
export type {
  DocxStyleResolveContext,
  DocxResolvedParagraphStyle,
  DocxResolvedRunStyle,
  DocxResolvedTableStyle,
  DocxResolvedNumberingLevel,
  DocxStyleIndex
};

/** @stability stable */
import type { CompatibilityMode as DocxCompatibilityMode } from "@word/query/compat";
export type { DocxCompatibilityMode };

/** @stability stable */
import type { FormFieldEntry as DocxFormFieldEntry } from "@word/query/form-fields";
export type { DocxFormFieldEntry };

/** @stability experimental */
import type { FieldEntry as DocxFieldEntry } from "@word/query/fields";
export type { DocxFieldEntry };

/** @stability stable */
import type {
  PatchContent as DocxPatchContent,
  PatchOperation as DocxPatchOperation,
  PatchOptions as DocxPatchOptions,
  CompiledTemplate as DocxCompiledTemplate,
  ToFlatOpcOptions as DocxToFlatOpcOptions
} from "@word/document-io";
export type {
  DocxPatchContent,
  DocxPatchOperation,
  DocxPatchOptions,
  DocxCompiledTemplate,
  DocxToFlatOpcOptions
};

// Incremental edit API (low-level efficient editing)
import type {
  IncrementalEdit as DocxIncrementalEdit,
  IncrementalEditOptions as DocxIncrementalEditOptions
} from "@word/incremental-edit";
export type { DocxIncrementalEdit, DocxIncrementalEditOptions };

// --- Stable API ---

/** @stability stable */
import type { ReadDocxOptions as DocxReadOptions } from "@word/reader/docx-reader";
export type { DocxReadOptions };
/** @stability stable */
export type { DocxWriteOptions } from "@word/writer/docx-packager";

// --- Experimental API ---
/** @stability experimental */
import type {
  DiffChangeType as DocxDiffChangeType,
  DiffEntry as DocxDiffEntry,
  DiffSummary as DocxDiffSummary,
  DiffResult as DocxDiffResult
} from "@word/advanced/diff";
export type { DocxDiffChangeType, DocxDiffEntry, DocxDiffSummary, DocxDiffResult };

/** @stability experimental */
import type {
  TemplateOptions as DocxTemplateOptions,
  TemplateImage as DocxTemplateImage,
  TemplateRichText as DocxTemplateRichText,
  TemplateSubDocument as DocxTemplateSubDocument,
  TemplateChart as DocxTemplateChart,
  TemplateHtmlChunk as DocxTemplateHtmlChunk,
  TemplateTag as DocxTemplateTag,
  LegacyTemplateOptions as DocxLegacyTemplateOptionsType
} from "@word/template/template-engine";
export type {
  DocxTemplateOptions,
  DocxTemplateImage,
  DocxTemplateRichText,
  DocxTemplateSubDocument,
  DocxTemplateChart,
  DocxTemplateHtmlChunk,
  DocxTemplateTag
};
/**
 * @deprecated The options of the deprecated template entry points (`Template.fillTemplate`,
 * `fillTemplateEnhanced` and their `…FromSource` / `…FromBuffer` forms); use
 * {@link DocxTemplateOptions} with `Template.fill`. Removed in 1.0.
 */
export type DocxLegacyTemplateOptions = DocxLegacyTemplateOptionsType;
/** @stability experimental */
import type {
  DataSource as DocxDataSource,
  FillFromSourceOptions as DocxFillFromSourceOptions
} from "@word/template/template-datasource";
export type { DocxFillFromSourceOptions };

// --- Stable API ---
/** @stability stable */
import type { FlatOpcPart as DocxFlatOpcPart } from "@word/convert/flat-opc";
export type { DocxFlatOpcPart };

// --- Experimental API ---
/** @stability experimental */
import type {
  StreamingDocxOptions as DocxStreamingWriteOptions,
  StreamingProgressCallback as DocxStreamingProgressCallback
} from "@word/writer/streaming-writer";
export type { DocxStreamingWriteOptions, DocxStreamingProgressCallback };

/** @stability experimental */
import type {
  StreamingReadOptions as DocxStreamingReadOptions,
  StreamingDocxMetadata as DocxStreamingMetadata
} from "@word/reader/streaming-reader";
export type { DocxStreamingReadOptions, DocxStreamingMetadata };

/** @stability experimental */
import type {
  ValidationSeverity as DocxValidationSeverity,
  ValidationIssue as DocxValidationIssue,
  ValidationResult as DocxValidationResult,
  ValidationOptions as DocxValidationOptions
} from "@word/advanced/validation";
export type {
  DocxValidationSeverity,
  DocxValidationIssue,
  DocxValidationResult,
  DocxValidationOptions
};

// Font embedding
/** @stability stable */
import type {
  FontEmbedStyle as DocxFontEmbedStyle,
  EmbedFontOptions as DocxEmbedFontOptions,
  EmbedFontResult as DocxEmbedFontResult
} from "@word/font/font-embed";
export type { DocxFontEmbedStyle, DocxEmbedFontOptions, DocxEmbedFontResult };

// --- Experimental API ---
/** @stability experimental */
import type {
  OleObject as DocxOleObject,
  OleObjectType as DocxOleObjectType,
  OleDisplayAs as DocxOleDisplayAs,
  OleExtractionResult as DocxOleExtractionResult,
  OleEmbeddingResult as DocxOleEmbeddingResult
} from "@word/advanced/ole-objects";
export type {
  DocxOleObject,
  DocxOleObjectType,
  DocxOleDisplayAs,
  DocxOleExtractionResult,
  DocxOleEmbeddingResult
};

/** @stability experimental */
import type {
  BuildingBlockGallery as DocxBuildingBlockGallery,
  BuildingBlock as DocxBuildingBlock,
  GlossaryDocument as DocxGlossaryDocument
} from "@word/advanced/glossary";
export type { DocxBuildingBlockGallery, DocxBuildingBlock, DocxGlossaryDocument };

// Document protection
/** @stability stable */
import type {
  ProtectionEditType as DocxProtectionEditType,
  ProtectionHashAlgorithm as DocxProtectionHashAlgorithm,
  DocumentProtectionOptions as DocxProtectionOptions,
  ProtectionState as DocxProtectionState
} from "@word/security/document-protection";
export type {
  DocxProtectionEditType,
  DocxProtectionHashAlgorithm,
  DocxProtectionOptions,
  DocxProtectionState
};

// Style mapping DSL
/** @stability stable */
import type {
  MappingSourceType as DocxMappingSourceType,
  MappingCondition as DocxMappingCondition,
  MappingTarget as DocxMappingTarget,
  StyleMappingRule as DocxStyleMappingRule,
  StyleMap as DocxStyleMap,
  StyleMapOptions as DocxStyleMapOptions
} from "@word/advanced/style-map";
export type {
  DocxMappingSourceType,
  DocxMappingCondition,
  DocxMappingTarget,
  DocxStyleMappingRule,
  DocxStyleMap,
  DocxStyleMapOptions
};

// --- Experimental API ---
/** @stability experimental */
import type { VbaProjectInfo as DocxVbaProjectInfo } from "@word/advanced/vba-project";
export type { DocxVbaProjectInfo };

/** @stability experimental */
import type {
  LayoutResult as DocxLayoutResult,
  LayoutOptions as DocxLayoutOptions
} from "@word/layout/layout";
export type { DocxLayoutResult, DocxLayoutOptions };
/** @stability experimental */
import type {
  FullLayoutOptions as DocxFullLayoutOptions,
  PageGeometryOverride as DocxPageGeometryOverride
} from "@word/layout/layout-full";
export type { DocxFullLayoutOptions, DocxPageGeometryOverride };

/** @stability experimental */
import type { FieldUpdateOptions as DocxFieldUpdateOptions } from "@word/advanced/field-engine";
export type { DocxFieldUpdateOptions };

/** @stability experimental */
import type {
  StandardShapeType as DocxStandardShapeType,
  SolidFill as DocxSolidFill,
  GradientStop as DocxGradientStop,
  GradientFill as DocxGradientFill,
  PatternFill as DocxPatternFill,
  NoFill as DocxNoFill,
  ShapeFill as DocxShapeFill,
  LineDash as DocxLineDash,
  LineEndType as DocxLineEndType,
  LineEndSize as DocxLineEndSize,
  ShapeOutline as DocxShapeOutline,
  ShadowEffect as DocxShadowEffect,
  GlowEffect as DocxGlowEffect,
  ReflectionEffect as DocxReflectionEffect,
  Effect3D as DocxEffect3D,
  ShapeEffects as DocxShapeEffects,
  TextVerticalAnchor as DocxTextVerticalAnchor,
  TextWrap as DocxTextWrap,
  ShapeTextBody as DocxShapeTextBody,
  CreateShapeOptions as DocxCreateShapeOptions
} from "@word/advanced/drawing-shapes";
export type {
  DocxStandardShapeType,
  DocxSolidFill,
  DocxGradientStop,
  DocxGradientFill,
  DocxPatternFill,
  DocxNoFill,
  DocxShapeFill,
  DocxLineDash,
  DocxLineEndType,
  DocxLineEndSize,
  DocxShapeOutline,
  DocxShadowEffect,
  DocxGlowEffect,
  DocxReflectionEffect,
  DocxEffect3D,
  DocxShapeEffects,
  DocxTextVerticalAnchor,
  DocxTextWrap,
  DocxShapeTextBody,
  DocxCreateShapeOptions
};

import type {
  ScriptType as DocxScriptType,
  BiDiDirection as DocxBiDiDirection,
  ShapedCluster as DocxShapedCluster,
  ShapingOptions as DocxShapingOptions
} from "@word/font/text-shaping";
export type { DocxScriptType, DocxBiDiDirection, DocxShapedCluster, DocxShapingOptions };

import type {
  HyphenateOptions as DocxHyphenateOptions,
  HyphenationOptions as DocxHyphenationOptions,
  HyphenationPatterns as DocxHyphenationPatterns
} from "@word/font/hyphenation";
export type { DocxHyphenateOptions, DocxHyphenationOptions, DocxHyphenationPatterns };

import type {
  FormatCriteria as DocxFormatCriteria,
  FormatSearchResult as DocxFormatSearchResult
} from "@word/query/format-search";
export type { DocxFormatCriteria, DocxFormatSearchResult };

import type {
  RenderChartContext as DocxRenderChartContext,
  RenderOptions as DocxRenderOptions
} from "@word/layout/render-page";
export type { DocxRenderChartContext, DocxRenderOptions };

export type { DocxDataSource };

import type {
  ChartBinding as DocxChartBinding,
  ChartSeriesData as DocxChartSeriesData,
  ChartTemplateData as DocxChartTemplateData
} from "@word/template/template-chart";
export type { DocxChartBinding, DocxChartSeriesData, DocxChartTemplateData };

// NOTE: the HTML, Markdown, Excel and encryption converters are NOT exported
// from the main entry — they live in dedicated subpaths to ensure tree-shaking:
//   import { WordMarkdown } from "documonster/word/markdown"
//   import { WordHtml }     from "documonster/word/html"
//   import { WordExcel }    from "documonster/word/excel"
//   import { WordCrypto }   from "documonster/word/crypto"
// DOCX → PDF lives in the PDF module's lazy converter namespace:
//   import { Pdf } from "documonster/pdf"   // await Pdf.fromDocx(doc)

// --- Stable API ---
import type { EncryptOptions as DocxEncryptOptions } from "@word/security/encryption";
export type { DocxEncryptOptions };

// --- Security Policy ---
/** @stability experimental */
import type { DocxSecurityPolicy } from "@word/security/policy";
export type { DocxSecurityPolicy };

// --- Render Context (experimental) ---
/** @stability experimental */
import type {
  DocxRenderContext,
  IdGenerators as DocxIdGenerators,
  PreservedLinkRId as DocxPreservedLinkRId
} from "@word/writer/render-context";
export type { DocxRenderContext, DocxIdGenerators, DocxPreservedLinkRId };

// --- Document Transformer (experimental) ---
/** @stability experimental */
import type { DocxTransformer, MapOptions as DocxMapOptions } from "@word/core/mapper";
export type { DocxTransformer, DocxMapOptions };

// --- Document Walker (experimental) ---
/** @stability experimental */
import type {
  DocxVisitor,
  WalkPath as DocxWalkPath,
  WalkOptions as DocxWalkOptions,
  VisitAction as DocxVisitAction
} from "@word/core/walker";
export type { DocxVisitor, DocxWalkPath, DocxWalkOptions, DocxVisitAction };

// --- Layout Model (experimental) ---
/** @stability experimental */
import type {
  LayoutDocument as DocxLayoutDocument,
  LayoutPage as DocxLayoutPage,
  PageGeometry as DocxPageGeometry,
  PageContent as DocxPageContent,
  LayoutParagraph as DocxLayoutParagraph,
  LayoutTable as DocxLayoutTable,
  LayoutTableCell as DocxLayoutTableCell,
  LayoutImage as DocxLayoutImage,
  LayoutFloat as DocxLayoutFloat,
  LayoutFloatWrap as DocxLayoutFloatWrap,
  LayoutTextBox as DocxLayoutTextBox,
  LayoutShape as DocxLayoutShape,
  LayoutChart as DocxLayoutChart,
  LayoutSdt as DocxLayoutSdt,
  LayoutMath as DocxLayoutMath,
  LayoutCheckBox as DocxLayoutCheckBox,
  LayoutTableOfContents as DocxLayoutTableOfContents,
  LayoutAltChunk as DocxLayoutAltChunk,
  LayoutOpaqueDrawing as DocxLayoutOpaqueDrawing,
  LineBox as DocxLineBox,
  LineBoxItem as DocxLineBoxItem,
  PositionedInlineImage as DocxPositionedInlineImage,
  PositionedRun as DocxPositionedRun
} from "@word/layout/layout-model";
export type {
  DocxLayoutDocument,
  DocxLayoutPage,
  DocxPageGeometry,
  DocxPageContent,
  DocxLayoutParagraph,
  DocxLayoutTable,
  DocxLayoutTableCell,
  DocxLayoutImage,
  DocxLayoutFloat,
  DocxLayoutFloatWrap,
  DocxLayoutTextBox,
  DocxLayoutShape,
  DocxLayoutChart,
  DocxLayoutSdt,
  DocxLayoutMath,
  DocxLayoutCheckBox,
  DocxLayoutTableOfContents,
  DocxLayoutAltChunk,
  DocxLayoutOpaqueDrawing,
  DocxLineBox,
  DocxLineBoxItem,
  DocxPositionedInlineImage,
  DocxPositionedRun
};

// --- Conversion IR (experimental) ---
/** @stability experimental */
import type {
  SemanticDocument as DocxSemanticDocument,
  SemanticBlock as DocxSemanticBlock,
  SemanticInline as DocxSemanticInline,
  ConversionContext as DocxConversionContext,
  ConversionWarning as DocxConversionWarning,
  ConversionAsset as DocxConversionAsset,
  ResolvedFormatting as DocxResolvedFormatting
} from "@word/convert/conversion-ir";
export type {
  DocxSemanticDocument,
  DocxSemanticBlock,
  DocxSemanticInline,
  DocxConversionContext,
  DocxConversionWarning,
  DocxConversionAsset,
  DocxResolvedFormatting
};
/** @stability experimental */
export type { DocxToSemanticOptions } from "@word/convert/docx-to-semantic";

// --- Model helpers ---
/** @stability experimental */
export type { AssetRef, MediaAsset } from "@word/types";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";

// =============================================================================
// Deprecated unprefixed type names — every flat operation type of this entry carries `Docx`
// =============================================================================

/** @deprecated Use {@link DocxSearchResult}. Removed in 1.0. */
export type SearchResult = DocxSearchResult;
/** @deprecated Use {@link DocxHeading}. Removed in 1.0. */
export type DocumentHeading = DocxHeading;
/** @deprecated Use {@link DocxSectionInfo}. Removed in 1.0. */
export type DocumentSection = DocxSectionInfo;
/** @deprecated Use {@link DocxBookmarkLocation}. Removed in 1.0. */
export type BookmarkLocation = DocxBookmarkLocation;
/** @deprecated Use {@link DocxListTablesOptions}. Removed in 1.0. */
export type ListTablesOptions = DocxListTablesOptions;
/** @deprecated Use {@link DocxMergeOptions}. Removed in 1.0. */
export type MergeOptions = DocxMergeOptions;
/** @deprecated Use {@link DocxSplitOptions}. Removed in 1.0. */
export type SplitOptions = DocxSplitOptions;
/** @deprecated Use {@link DocxRevisionEntry}. Removed in 1.0. */
export type RevisionEntry = DocxRevisionEntry;
/** @deprecated Use {@link DocxStyleResolveContext}. Removed in 1.0. */
export type StyleResolveContext = DocxStyleResolveContext;
/** @deprecated Use {@link DocxResolvedParagraphStyle}. Removed in 1.0. */
export type ResolvedParagraphStyle = DocxResolvedParagraphStyle;
/** @deprecated Use {@link DocxResolvedRunStyle}. Removed in 1.0. */
export type ResolvedRunStyle = DocxResolvedRunStyle;
/** @deprecated Use {@link DocxResolvedTableStyle}. Removed in 1.0. */
export type ResolvedTableStyle = DocxResolvedTableStyle;
/** @deprecated Use {@link DocxResolvedNumberingLevel}. Removed in 1.0. */
export type ResolvedNumberingLevel = DocxResolvedNumberingLevel;
/** @deprecated Use {@link DocxStyleIndex}. Removed in 1.0. */
export type StyleIndex = DocxStyleIndex;
/** @deprecated Use {@link DocxCompatibilityMode}. Removed in 1.0. */
export type CompatibilityMode = DocxCompatibilityMode;
/** @deprecated Use {@link DocxFormFieldEntry}. Removed in 1.0. */
export type FormFieldEntry = DocxFormFieldEntry;
/** @deprecated Use {@link DocxFieldEntry}. Removed in 1.0. */
export type FieldEntry = DocxFieldEntry;
/** @deprecated Use {@link DocxPatchContent}. Removed in 1.0. */
export type PatchContent = DocxPatchContent;
/** @deprecated Use {@link DocxPatchOperation}. Removed in 1.0. */
export type PatchOperation = DocxPatchOperation;
/** @deprecated Use {@link DocxPatchOptions}. Removed in 1.0. */
export type PatchOptions = DocxPatchOptions;
/** @deprecated Use {@link DocxCompiledTemplate}. Removed in 1.0. */
export type CompiledTemplate = DocxCompiledTemplate;
/** @deprecated Use {@link DocxToFlatOpcOptions}. Removed in 1.0. */
export type ToFlatOpcOptions = DocxToFlatOpcOptions;
/** @deprecated Use {@link DocxIncrementalEdit}. Removed in 1.0. */
export type IncrementalEdit = DocxIncrementalEdit;
/** @deprecated Use {@link DocxIncrementalEditOptions}. Removed in 1.0. */
export type IncrementalEditOptions = DocxIncrementalEditOptions;
/** @deprecated Use {@link DocxReadOptions}. Removed in 1.0. */
export type ReadDocxOptions = DocxReadOptions;
/** @deprecated Use {@link DocxDiffChangeType}. Removed in 1.0. */
export type DiffChangeType = DocxDiffChangeType;
/** @deprecated Use {@link DocxDiffEntry}. Removed in 1.0. */
export type DiffEntry = DocxDiffEntry;
/** @deprecated Use {@link DocxDiffSummary}. Removed in 1.0. */
export type DiffSummary = DocxDiffSummary;
/** @deprecated Use {@link DocxDiffResult}. Removed in 1.0. */
export type DiffResult = DocxDiffResult;
/** @deprecated Use {@link DocxTemplateOptions}. Removed in 1.0. */
export type TemplateOptions = DocxTemplateOptions;
/** @deprecated Use {@link DocxTemplateImage}. Removed in 1.0. */
export type TemplateImage = DocxTemplateImage;
/** @deprecated Use {@link DocxTemplateRichText}. Removed in 1.0. */
export type TemplateRichText = DocxTemplateRichText;
/** @deprecated Use {@link DocxTemplateSubDocument}. Removed in 1.0. */
export type TemplateSubDocument = DocxTemplateSubDocument;
/** @deprecated Use {@link DocxTemplateChart}. Removed in 1.0. */
export type TemplateChart = DocxTemplateChart;
/** @deprecated Use {@link DocxTemplateHtmlChunk}. Removed in 1.0. */
export type TemplateHtmlChunk = DocxTemplateHtmlChunk;
/** @deprecated Use {@link DocxTemplateTag}. Removed in 1.0. */
export type TemplateTag = DocxTemplateTag;
/** @deprecated Use {@link DocxTemplateOptions} with `Template.fill`. Removed in 1.0. */
export type LegacyTemplateOptions = DocxLegacyTemplateOptionsType;
/** @deprecated Use {@link DocxFillFromSourceOptions}. Removed in 1.0. */
export type FillFromSourceOptions = DocxFillFromSourceOptions;
/** @deprecated Use {@link DocxFlatOpcPart}. Removed in 1.0. */
export type FlatOpcPart = DocxFlatOpcPart;
/** @deprecated Use {@link DocxStreamingWriteOptions}. Removed in 1.0. */
export type StreamingDocxOptions = DocxStreamingWriteOptions;
/** @deprecated Use {@link DocxStreamingProgressCallback}. Removed in 1.0. */
export type StreamingProgressCallback = DocxStreamingProgressCallback;
/** @deprecated Use {@link DocxStreamingReadOptions}. Removed in 1.0. */
export type StreamingReadOptions = DocxStreamingReadOptions;
/** @deprecated Use {@link DocxStreamingMetadata}. Removed in 1.0. */
export type StreamingDocxMetadata = DocxStreamingMetadata;
/** @deprecated Use {@link DocxValidationSeverity}. Removed in 1.0. */
export type ValidationSeverity = DocxValidationSeverity;
/** @deprecated Use {@link DocxValidationIssue}. Removed in 1.0. */
export type ValidationIssue = DocxValidationIssue;
/** @deprecated Use {@link DocxValidationResult}. Removed in 1.0. */
export type ValidationResult = DocxValidationResult;
/** @deprecated Use {@link DocxValidationOptions}. Removed in 1.0. */
export type ValidationOptions = DocxValidationOptions;
/** @deprecated Use {@link DocxFontEmbedStyle}. Removed in 1.0. */
export type FontEmbedStyle = DocxFontEmbedStyle;
/** @deprecated Use {@link DocxEmbedFontOptions}. Removed in 1.0. */
export type EmbedFontOptions = DocxEmbedFontOptions;
/** @deprecated Use {@link DocxEmbedFontResult}. Removed in 1.0. */
export type EmbedFontResult = DocxEmbedFontResult;
/** @deprecated Use {@link DocxOleObject}. Removed in 1.0. */
export type OleObject = DocxOleObject;
/** @deprecated Use {@link DocxOleObjectType}. Removed in 1.0. */
export type OleObjectType = DocxOleObjectType;
/** @deprecated Use {@link DocxOleDisplayAs}. Removed in 1.0. */
export type OleDisplayAs = DocxOleDisplayAs;
/** @deprecated Use {@link DocxOleExtractionResult}. Removed in 1.0. */
export type OleExtractionResult = DocxOleExtractionResult;
/** @deprecated Use {@link DocxOleEmbeddingResult}. Removed in 1.0. */
export type OleEmbeddingResult = DocxOleEmbeddingResult;
/** @deprecated Use {@link DocxBuildingBlockGallery}. Removed in 1.0. */
export type BuildingBlockGallery = DocxBuildingBlockGallery;
/** @deprecated Use {@link DocxBuildingBlock}. Removed in 1.0. */
export type BuildingBlock = DocxBuildingBlock;
/** @deprecated Use {@link DocxGlossaryDocument}. Removed in 1.0. */
export type GlossaryDocument = DocxGlossaryDocument;
/** @deprecated Use {@link DocxProtectionEditType}. Removed in 1.0. */
export type ProtectionEditType = DocxProtectionEditType;
/** @deprecated Use {@link DocxProtectionHashAlgorithm}. Removed in 1.0. */
export type ProtectionHashAlgorithm = DocxProtectionHashAlgorithm;
/** @deprecated Use {@link DocxProtectionOptions}. Removed in 1.0. */
export type DocumentProtectionOptions = DocxProtectionOptions;
/** @deprecated Use {@link DocxProtectionState}. Removed in 1.0. */
export type ProtectionState = DocxProtectionState;
/** @deprecated Use {@link DocxMappingSourceType}. Removed in 1.0. */
export type MappingSourceType = DocxMappingSourceType;
/** @deprecated Use {@link DocxMappingCondition}. Removed in 1.0. */
export type MappingCondition = DocxMappingCondition;
/** @deprecated Use {@link DocxMappingTarget}. Removed in 1.0. */
export type MappingTarget = DocxMappingTarget;
/** @deprecated Use {@link DocxStyleMappingRule}. Removed in 1.0. */
export type StyleMappingRule = DocxStyleMappingRule;
/** @deprecated Use {@link DocxStyleMap}. Removed in 1.0. */
export type StyleMap = DocxStyleMap;
/** @deprecated Use {@link DocxStyleMapOptions}. Removed in 1.0. */
export type StyleMapOptions = DocxStyleMapOptions;
/** @deprecated Use {@link DocxVbaProjectInfo}. Removed in 1.0. */
export type VbaProjectInfo = DocxVbaProjectInfo;
/** @deprecated Use {@link DocxLayoutResult}. Removed in 1.0. */
export type LayoutResult = DocxLayoutResult;
/** @deprecated Use {@link DocxLayoutOptions}. Removed in 1.0. */
export type LayoutOptions = DocxLayoutOptions;
/** @deprecated Use {@link DocxFullLayoutOptions}. Removed in 1.0. */
export type FullLayoutOptions = DocxFullLayoutOptions;
/** @deprecated Use {@link DocxPageGeometryOverride}. Removed in 1.0. */
export type PageGeometryOverride = DocxPageGeometryOverride;
/** @deprecated Use {@link DocxFieldUpdateOptions}. Removed in 1.0. */
export type FieldUpdateOptions = DocxFieldUpdateOptions;
/** @deprecated Use {@link DocxStandardShapeType}. Removed in 1.0. */
export type StandardShapeType = DocxStandardShapeType;
/** @deprecated Use {@link DocxSolidFill}. Removed in 1.0. */
export type SolidFill = DocxSolidFill;
/** @deprecated Use {@link DocxGradientStop}. Removed in 1.0. */
export type GradientStop = DocxGradientStop;
/** @deprecated Use {@link DocxGradientFill}. Removed in 1.0. */
export type GradientFill = DocxGradientFill;
/** @deprecated Use {@link DocxPatternFill}. Removed in 1.0. */
export type PatternFill = DocxPatternFill;
/** @deprecated Use {@link DocxNoFill}. Removed in 1.0. */
export type NoFill = DocxNoFill;
/** @deprecated Use {@link DocxShapeFill}. Removed in 1.0. */
export type ShapeFill = DocxShapeFill;
/** @deprecated Use {@link DocxLineDash}. Removed in 1.0. */
export type LineDash = DocxLineDash;
/** @deprecated Use {@link DocxLineEndType}. Removed in 1.0. */
export type LineEndType = DocxLineEndType;
/** @deprecated Use {@link DocxLineEndSize}. Removed in 1.0. */
export type LineEndSize = DocxLineEndSize;
/** @deprecated Use {@link DocxShapeOutline}. Removed in 1.0. */
export type ShapeOutline = DocxShapeOutline;
/** @deprecated Use {@link DocxShadowEffect}. Removed in 1.0. */
export type ShadowEffect = DocxShadowEffect;
/** @deprecated Use {@link DocxGlowEffect}. Removed in 1.0. */
export type GlowEffect = DocxGlowEffect;
/** @deprecated Use {@link DocxReflectionEffect}. Removed in 1.0. */
export type ReflectionEffect = DocxReflectionEffect;
/** @deprecated Use {@link DocxEffect3D}. Removed in 1.0. */
export type Effect3D = DocxEffect3D;
/** @deprecated Use {@link DocxShapeEffects}. Removed in 1.0. */
export type ShapeEffects = DocxShapeEffects;
/** @deprecated Use {@link DocxTextVerticalAnchor}. Removed in 1.0. */
export type TextVerticalAnchor = DocxTextVerticalAnchor;
/** @deprecated Use {@link DocxTextWrap}. Removed in 1.0. */
export type TextWrap = DocxTextWrap;
/** @deprecated Use {@link DocxShapeTextBody}. Removed in 1.0. */
export type ShapeTextBody = DocxShapeTextBody;
/** @deprecated Use {@link DocxCreateShapeOptions}. Removed in 1.0. */
export type CreateShapeOptions = DocxCreateShapeOptions;
/** @deprecated Use {@link DocxScriptType}. Removed in 1.0. */
export type ScriptType = DocxScriptType;
/** @deprecated Use {@link DocxBiDiDirection}. Removed in 1.0. */
export type BiDiDirection = DocxBiDiDirection;
/** @deprecated Use {@link DocxShapedCluster}. Removed in 1.0. */
export type ShapedCluster = DocxShapedCluster;
/** @deprecated Use {@link DocxShapingOptions}. Removed in 1.0. */
export type ShapingOptions = DocxShapingOptions;
/** @deprecated Use {@link DocxHyphenationOptions}. Removed in 1.0. */
export type HyphenationOptions = DocxHyphenationOptions;
/** @deprecated Use {@link DocxHyphenationPatterns}. Removed in 1.0. */
export type HyphenationPatterns = DocxHyphenationPatterns;
/** @deprecated Use {@link DocxFormatCriteria}. Removed in 1.0. */
export type FormatCriteria = DocxFormatCriteria;
/** @deprecated Use {@link DocxFormatSearchResult}. Removed in 1.0. */
export type FormatSearchResult = DocxFormatSearchResult;
/** @deprecated Use {@link DocxRenderChartContext}. Removed in 1.0. */
export type RenderChartContext = DocxRenderChartContext;
/** @deprecated Use {@link DocxRenderOptions}. Removed in 1.0. */
export type RenderOptions = DocxRenderOptions;
/** @deprecated Use {@link DocxDataSource}. Removed in 1.0. */
export type DataSource = DocxDataSource;
/** @deprecated Use {@link DocxChartBinding}. Removed in 1.0. */
export type ChartBinding = DocxChartBinding;
/** @deprecated Use {@link DocxChartSeriesData}. Removed in 1.0. */
export type ChartSeriesData = DocxChartSeriesData;
/** @deprecated Use {@link DocxChartTemplateData}. Removed in 1.0. */
export type ChartTemplateData = DocxChartTemplateData;
/** @deprecated Use {@link DocxEncryptOptions}. Removed in 1.0. */
export type EncryptOptions = DocxEncryptOptions;
/** @deprecated Use {@link DocxSecurityPolicy}. Removed in 1.0. */
export type WordSecurityPolicy = DocxSecurityPolicy;
/** @deprecated Use {@link DocxRenderContext}. Removed in 1.0. */
export type WordRenderContext = DocxRenderContext;
/** @deprecated Use {@link DocxIdGenerators}. Removed in 1.0. */
export type IdGenerators = DocxIdGenerators;
/** @deprecated Use {@link DocxPreservedLinkRId}. Removed in 1.0. */
export type PreservedLinkRId = DocxPreservedLinkRId;
/** @deprecated Use {@link DocxMapOptions}. Removed in 1.0. */
export type MapOptions = DocxMapOptions;
/** @deprecated Use {@link DocxWalkPath}. Removed in 1.0. */
export type WalkPath = DocxWalkPath;
/** @deprecated Use {@link DocxWalkOptions}. Removed in 1.0. */
export type WalkOptions = DocxWalkOptions;
/** @deprecated Use {@link DocxVisitAction}. Removed in 1.0. */
export type VisitAction = DocxVisitAction;
/** @deprecated Use {@link DocxLayoutDocument}. Removed in 1.0. */
export type LayoutDocument = DocxLayoutDocument;
/** @deprecated Use {@link DocxLayoutPage}. Removed in 1.0. */
export type LayoutPage = DocxLayoutPage;
/** @deprecated Use {@link DocxPageGeometry}. Removed in 1.0. */
export type PageGeometry = DocxPageGeometry;
/** @deprecated Use {@link DocxPageContent}. Removed in 1.0. */
export type PageContent = DocxPageContent;
/** @deprecated Use {@link DocxLayoutParagraph}. Removed in 1.0. */
export type LayoutParagraph = DocxLayoutParagraph;
/** @deprecated Use {@link DocxLayoutTable}. Removed in 1.0. */
export type LayoutTable = DocxLayoutTable;
/** @deprecated Use {@link DocxLayoutTableCell}. Removed in 1.0. */
export type LayoutTableCell = DocxLayoutTableCell;
/** @deprecated Use {@link DocxLayoutImage}. Removed in 1.0. */
export type LayoutImage = DocxLayoutImage;
/** @deprecated Use {@link DocxLayoutFloat}. Removed in 1.0. */
export type LayoutFloat = DocxLayoutFloat;
/** @deprecated Use {@link DocxLayoutFloatWrap}. Removed in 1.0. */
export type LayoutFloatWrap = DocxLayoutFloatWrap;
/** @deprecated Use {@link DocxLayoutTextBox}. Removed in 1.0. */
export type LayoutTextBox = DocxLayoutTextBox;
/** @deprecated Use {@link DocxLayoutShape}. Removed in 1.0. */
export type LayoutShape = DocxLayoutShape;
/** @deprecated Use {@link DocxLayoutChart}. Removed in 1.0. */
export type LayoutChart = DocxLayoutChart;
/** @deprecated Use {@link DocxLayoutSdt}. Removed in 1.0. */
export type LayoutSdt = DocxLayoutSdt;
/** @deprecated Use {@link DocxLayoutMath}. Removed in 1.0. */
export type LayoutMath = DocxLayoutMath;
/** @deprecated Use {@link DocxLayoutCheckBox}. Removed in 1.0. */
export type LayoutCheckBox = DocxLayoutCheckBox;
/** @deprecated Use {@link DocxLayoutTableOfContents}. Removed in 1.0. */
export type LayoutTableOfContents = DocxLayoutTableOfContents;
/** @deprecated Use {@link DocxLayoutAltChunk}. Removed in 1.0. */
export type LayoutAltChunk = DocxLayoutAltChunk;
/** @deprecated Use {@link DocxLayoutOpaqueDrawing}. Removed in 1.0. */
export type LayoutOpaqueDrawing = DocxLayoutOpaqueDrawing;
/** @deprecated Use {@link DocxLineBox}. Removed in 1.0. */
export type LineBox = DocxLineBox;
/** @deprecated Use {@link DocxLineBoxItem}. Removed in 1.0. */
export type LineBoxItem = DocxLineBoxItem;
/** @deprecated Use {@link DocxPositionedInlineImage}. Removed in 1.0. */
export type PositionedInlineImage = DocxPositionedInlineImage;
/** @deprecated Use {@link DocxPositionedRun}. Removed in 1.0. */
export type PositionedRun = DocxPositionedRun;
/** @deprecated Use {@link DocxSemanticDocument}. Removed in 1.0. */
export type SemanticDocument = DocxSemanticDocument;
/** @deprecated Use {@link DocxSemanticBlock}. Removed in 1.0. */
export type SemanticBlock = DocxSemanticBlock;
/** @deprecated Use {@link DocxSemanticInline}. Removed in 1.0. */
export type SemanticInline = DocxSemanticInline;
/** @deprecated Use {@link DocxConversionContext}. Removed in 1.0. */
export type ConversionContext = DocxConversionContext;
/** @deprecated Use {@link DocxConversionWarning}. Removed in 1.0. */
export type ConversionWarning = DocxConversionWarning;
/** @deprecated Use {@link DocxConversionAsset}. Removed in 1.0. */
export type ConversionAsset = DocxConversionAsset;
/** @deprecated Use {@link DocxResolvedFormatting}. Removed in 1.0. */
export type ResolvedFormatting = DocxResolvedFormatting;
