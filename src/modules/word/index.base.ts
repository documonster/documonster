/**
 * DOCX Module - Base Exports (Platform Independent)
 *
 * Shared exports for both Node.js and browser environments.
 */

// Types
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

  // Alt chunks (embedded HTML/RTF content)
  AltChunk,

  // Web settings & people (collaboration metadata)
  WebSettings,
  PersonInfo,

  // Notes type enum
  NoteType,

  // Table revision types
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  CellMergeRevision,

  // Opaque parts (round-trip preservation)
  OpaquePart,
  OpaqueRelationship,
  OpaqueDrawing,

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

// Errors
export {
  DocxError,
  DocxParseError,
  DocxWriteError,
  DocxMissingPartError,
  DocxInvalidStructureError,
  DocxUnsupportedFeatureError,
  isDocxError
} from "./errors";

// Units
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

// Builder helpers
export {
  DocumentBuilder,
  // Run helpers
  text,
  bold,
  italic,
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
  patchDocument,
  mailMerge,
  // Query API
  paragraphCount,
  countWords,
  getHeadings,
  findBookmark,
  findComment,
  listImages,
  listTables,
  listHyperlinks,
  tableCount,
  extractText
} from "./document";
export type {
  SearchResult,
  PatchContent,
  PatchOperation,
  PatchOptions,
  DocumentHeading
} from "./document";

// Packager
export { packageDocx } from "./docx-packager";

// Reader
export { readDocx } from "./docx-reader";

// Font obfuscation utilities
export { deobfuscateFont, obfuscateFont, generateFontKey } from "./font-obfuscation";

// Encryption utilities
export {
  isEncryptedDocx,
  verifyPassword,
  decryptPackage,
  parseEncryptionInfoXml,
  deriveEncryptionKey,
  AGILE_BLOCK_KEYS
} from "./encryption";
export type { AgileEncryptionInfo } from "./encryption";

// Digital signature utilities
export {
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature
} from "./digital-signatures";
export type { DigitalSignatureInfo } from "./digital-signatures";

// HTML Renderer
export { renderToHtml } from "./html-renderer";
export type { HtmlRenderOptions, HtmlRenderResult } from "./html-renderer";
