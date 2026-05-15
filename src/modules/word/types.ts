/**
 * DOCX Module - Core Types
 *
 * Comprehensive type definitions for WordprocessingML (OOXML) documents.
 * Covers paragraphs, runs, tables, images, headers/footers, numbering,
 * styles, sections, footnotes/endnotes, bookmarks, hyperlinks, comments,
 * track changes, TOC, math equations, text boxes, checkboxes, custom
 * properties, page borders, document background, and more.
 */

// =============================================================================
// Units & Measurement
// =============================================================================

/**
 * Half-point value (used for font sizes).
 * A value of 24 = 12pt.
 */
export type HalfPoint = number;

/**
 * Twips (twentieths of a point).
 * 1 inch = 1440 twips, 1 cm = 567 twips, 1 pt = 20 twips.
 */
export type Twips = number;

/**
 * English Metric Units (EMU).
 * 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU.
 */
export type Emu = number;

/**
 * Eighths of a point (used for border widths).
 * A value of 4 = 0.5pt, 8 = 1pt.
 */
export type EighthPoint = number;

// =============================================================================
// Color
// =============================================================================

/**
 * 6-digit hex RGB color string (e.g. "FF0000") or "auto" for default.
 */
export type HexColor = string;

// =============================================================================
// Shading Pattern (ST_Shd) — Full OOXML enumeration
// =============================================================================

/** Shading pattern type (ST_Shd). */
export type ShadingType =
  | "clear"
  | "solid"
  | "horzStripe"
  | "vertStripe"
  | "reverseDiagStripe"
  | "diagStripe"
  | "horzCross"
  | "diagCross"
  | "thinHorzStripe"
  | "thinVertStripe"
  | "thinReverseDiagStripe"
  | "thinDiagStripe"
  | "thinHorzCross"
  | "thinDiagCross"
  | "pct5"
  | "pct10"
  | "pct12"
  | "pct15"
  | "pct20"
  | "pct25"
  | "pct30"
  | "pct35"
  | "pct37"
  | "pct40"
  | "pct45"
  | "pct50"
  | "pct55"
  | "pct60"
  | "pct62"
  | "pct65"
  | "pct70"
  | "pct75"
  | "pct80"
  | "pct85"
  | "pct87"
  | "pct90"
  | "pct95"
  | "nil";

// =============================================================================
// Border Style (ST_Border) — Full OOXML enumeration
// =============================================================================

/** Border style type (ST_Border). */
export type BorderStyle =
  | "single"
  | "double"
  | "thick"
  | "dotted"
  | "dashed"
  | "dotDash"
  | "dotDotDash"
  | "triple"
  | "thinThickSmallGap"
  | "thickThinSmallGap"
  | "thinThickThinSmallGap"
  | "thinThickMediumGap"
  | "thickThinMediumGap"
  | "thinThickThinMediumGap"
  | "thinThickLargeGap"
  | "thickThinLargeGap"
  | "thinThickThinLargeGap"
  | "wave"
  | "doubleWave"
  | "dashSmallGap"
  | "dashDotStroked"
  | "threeDEmboss"
  | "threeDEngrave"
  | "outset"
  | "inset"
  | "none"
  | "nil";

// =============================================================================
// Page & Section
// =============================================================================

/** Page orientation. */
export type PageOrientation = "portrait" | "landscape";

/** Page size preset. */
export interface PageSize {
  /** Width in twips. */
  readonly width: Twips;
  /** Height in twips. */
  readonly height: Twips;
  /** Orientation. Default: "portrait". */
  readonly orientation?: PageOrientation;
}

/** Page margins in twips. */
export interface PageMargins {
  readonly top: Twips;
  readonly right: Twips;
  readonly bottom: Twips;
  readonly left: Twips;
  /** Header distance from top edge. Default: 720 (0.5 inch). */
  readonly header?: Twips;
  /** Footer distance from bottom edge. Default: 720 (0.5 inch). */
  readonly footer?: Twips;
  /** Gutter margin. Default: 0. */
  readonly gutter?: Twips;
}

/** Column definition for multi-column sections. */
export interface ColumnDef {
  /** Column width in twips. */
  readonly width: Twips;
  /** Space after this column in twips. */
  readonly space?: Twips;
}

/** Section column layout. */
export interface SectionColumns {
  /** Number of columns (when equal width). */
  readonly count?: number;
  /** Space between columns in twips (when equal width). */
  readonly space?: Twips;
  /** Whether columns have equal width. Default: true. */
  readonly equalWidth?: boolean;
  /** Individual column definitions (when not equal width). */
  readonly columns?: readonly ColumnDef[];
  /** Whether to draw a line between columns. */
  readonly separator?: boolean;
}

/** Section break type. */
export type SectionBreakType = "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";

/** Page number format (ST_NumberFormat subset used for page numbering). */
export type PageNumberFormat =
  | "decimal"
  | "upperRoman"
  | "lowerRoman"
  | "upperLetter"
  | "lowerLetter"
  | "ordinal"
  | "cardinalText"
  | "ordinalText"
  | "hex"
  | "chicago"
  | "ideographDigital"
  | "japaneseCounting"
  | "aiueo"
  | "iroha"
  | "decimalFullWidth"
  | "decimalHalfWidth"
  | "japaneseLegal"
  | "japaneseDigitalTenThousand"
  | "decimalEnclosedCircle"
  | "decimalFullWidth2"
  | "aiueoFullWidth"
  | "irohaFullWidth"
  | "decimalZero"
  | "bullet"
  | "ganada"
  | "chosung"
  | "decimalEnclosedFullstop"
  | "decimalEnclosedParen"
  | "decimalEnclosedCircleChinese"
  | "ideographEnclosedCircle"
  | "ideographTraditional"
  | "ideographZodiac"
  | "ideographZodiacTraditional"
  | "taiwaneseCounting"
  | "ideographLegalTraditional"
  | "taiwaneseCountingThousand"
  | "taiwaneseDigital"
  | "chineseCounting"
  | "chineseLegalSimplified"
  | "chineseCountingThousand"
  | "koreanDigital"
  | "koreanCounting"
  | "koreanLegal"
  | "koreanDigital2"
  | "vietnameseCounting"
  | "russianLower"
  | "russianUpper"
  | "none"
  | "numberInDash"
  | "hebrew1"
  | "hebrew2"
  | "arabicAlpha"
  | "arabicAbjad"
  | "hindiVowels"
  | "hindiConsonants"
  | "hindiNumbers"
  | "hindiCounting"
  | "thaiLetters"
  | "thaiNumbers"
  | "thaiCounting"
  | "bahtText"
  | "dollarText"
  | "custom";

/** Header/footer reference type. */
export type HeaderFooterType = "default" | "first" | "even";

/** Reference to a header or footer part. */
export interface HeaderFooterRef {
  /** Type of header/footer. */
  readonly type: HeaderFooterType;
  /** Relationship ID referencing the header/footer part. */
  readonly rId: string;
}

/** Page vertical alignment. */
export type PageVerticalAlign = "top" | "center" | "both" | "bottom";

/** Document grid type. */
export type DocumentGridType = "default" | "lines" | "linesAndChars" | "snapToChars";

/** A single border definition. */
export interface Border {
  /** Border style. */
  readonly style: BorderStyle;
  /** Width in eighths of a point. */
  readonly size?: EighthPoint;
  /** Space between border and content in points. */
  readonly space?: number;
  /** Color (hex RGB or "auto"). */
  readonly color?: HexColor;
  /** Theme color. */
  readonly themeColor?: string;
  /** Shadow. */
  readonly shadow?: boolean;
  /** Frame. */
  readonly frame?: boolean;
  /** Artistic border preset (for page borders only, w:art attribute). */
  readonly art?: ArtBorderType;
}

/** Artistic page border types. There are 160+ standard Office artwork borders. */
export type ArtBorderType = string;

/** Page borders. */
export interface PageBorders {
  readonly top?: Border;
  readonly left?: Border;
  readonly bottom?: Border;
  readonly right?: Border;
  /** Display setting: "allPages" | "firstPage" | "notFirstPage". */
  readonly display?: "allPages" | "firstPage" | "notFirstPage";
  /** Offset from: "page" | "text". */
  readonly offsetFrom?: "page" | "text";
  /** Z-ordering: "front" | "back". */
  readonly zOrder?: "front" | "back";
}

/** Section properties. */
export interface SectionProperties {
  /** Page size. */
  readonly pageSize?: PageSize;
  /** Page margins. */
  readonly margins?: PageMargins;
  /** Section break type. */
  readonly breakType?: SectionBreakType;
  /** Column layout. */
  readonly columns?: SectionColumns;
  /** Header references. */
  readonly headers?: readonly HeaderFooterRef[];
  /** Footer references. */
  readonly footers?: readonly HeaderFooterRef[];
  /** Different first page header/footer. */
  readonly titlePage?: boolean;
  /** Page number settings. */
  readonly pageNumbering?: {
    readonly start?: number;
    readonly format?: PageNumberFormat;
  };
  /** Document grid. */
  readonly docGrid?: {
    readonly linePitch?: Twips;
    readonly charSpace?: number;
    readonly type?: DocumentGridType;
  };
  /** Line numbering. */
  readonly lineNumbers?: {
    readonly countBy?: number;
    readonly start?: number;
    readonly restart?: "newPage" | "newSection" | "continuous";
    readonly distance?: Twips;
  };
  /** Page borders. */
  readonly pageBorders?: PageBorders;
  /** Page vertical alignment. */
  readonly verticalAlign?: PageVerticalAlign;
  /** Page text direction. */
  readonly textDirection?: TextDirection;
  /** Footnote properties for this section. */
  readonly footnoteProperties?: FootnoteProperties;
  /** Endnote properties for this section. */
  readonly endnoteProperties?: EndnoteProperties;
  /** Section property change revision (track changes). */
  readonly propertyChange?: SectionPropertyChange;
  /** Right-to-left section (w:bidi). */
  readonly bidi?: boolean;
  /** Form protection (w:formProt). */
  readonly formProtection?: boolean;
  /** Register (w:rtlGutter for right-to-left). */
  readonly rtlGutter?: boolean;
}

// =============================================================================
// Run Properties (Character Formatting)
// =============================================================================

/** Underline styles (ST_Underline). */
export type UnderlineStyle =
  | "single"
  | "words"
  | "double"
  | "thick"
  | "dotted"
  | "dottedHeavy"
  | "dash"
  | "dashedHeavy"
  | "dashLong"
  | "dashLongHeavy"
  | "dotDash"
  | "dashDotHeavy"
  | "dotDotDash"
  | "dashDotDotHeavy"
  | "wave"
  | "wavyHeavy"
  | "wavyDouble"
  | "none";

/** Vertical alignment for superscript/subscript. */
export type VerticalAlign = "superscript" | "subscript" | "baseline";

/** Highlight colors. */
export type HighlightColor =
  | "black"
  | "blue"
  | "cyan"
  | "darkBlue"
  | "darkCyan"
  | "darkGray"
  | "darkGreen"
  | "darkMagenta"
  | "darkRed"
  | "darkYellow"
  | "green"
  | "lightGray"
  | "magenta"
  | "none"
  | "red"
  | "white"
  | "yellow";

/** Font specification. */
export interface FontSpec {
  /** ASCII / Latin font name. */
  readonly ascii?: string;
  /** High ANSI font name (often same as ascii). */
  readonly hAnsi?: string;
  /** East Asian font name. */
  readonly eastAsia?: string;
  /** Complex Script font name. */
  readonly cs?: string;
  /** Font hint. */
  readonly hint?: "default" | "eastAsia" | "cs";
  /** ASCII theme font reference (e.g. "minorHAnsi"). */
  readonly asciiTheme?: string;
  /** High ANSI theme font reference. */
  readonly hAnsiTheme?: string;
  /** East Asian theme font reference. */
  readonly eastAsiaTheme?: string;
  /** Complex Script theme font reference. */
  readonly cstheme?: string;
}

/** Shading / background fill. */
export interface Shading {
  /** Shading pattern. Default: "clear". */
  readonly pattern?: ShadingType;
  /** Pattern color. Usually "auto". */
  readonly color?: HexColor;
  /** Fill / background color. */
  readonly fill: HexColor;
}

/** Underline specification (style + optional color). */
export interface UnderlineSpec {
  /** Underline style. */
  readonly style: UnderlineStyle;
  /** Underline color (hex RGB or "auto"). Independent of text color. */
  readonly color?: HexColor;
}

/** Text effect animation. */
export type TextEffect =
  | "blinkBackground"
  | "lights"
  | "antsBlack"
  | "antsRed"
  | "shimmer"
  | "sparkle"
  | "none";

/** Emphasis mark type. */
export type EmphasisMarkType = "dot" | "comma" | "circle" | "underDot" | "none";

/** Bracket style used when combining East Asian characters into a composite cell. */
export type EastAsianCombineBrackets = "none" | "round" | "square" | "angle" | "curly";

/**
 * East Asian typographic layout overrides for a run (ECMA-376 §17.3.2.10
 * `<w:eastAsianLayout>`). Used for Japanese/Korean classical layout to
 * combine consecutive characters into a single display cell, render
 * vertically, or compress vertical glyphs.
 *
 * The layout/render-page modules render runs carrying this property in
 * standard horizontal direction (vert/vertCompress are not visualised);
 * reader and writer round-trip the property losslessly so the source
 * intent is preserved across edits.
 */
export interface EastAsianLayoutSpec {
  /** Tracking id; pairs with `rPrChange` revisions. */
  readonly id?: number;
  /** Combine consecutive characters into one composite display cell. */
  readonly combine?: boolean;
  /** Bracket style used when combining (defaults to `none`). */
  readonly combineBrackets?: EastAsianCombineBrackets;
  /** Render the run vertically (top-to-bottom). */
  readonly vert?: boolean;
  /** Compress vertical glyphs into half-width cells. */
  readonly vertCompress?: boolean;
}

/** Color specification with optional theme support. */
export interface ColorSpec {
  /** Hex RGB color value (e.g. "FF0000") or "auto". */
  readonly val: HexColor;
  /** Theme color reference. */
  readonly themeColor?: string;
  /** Theme tint (hex, e.g. "99"). */
  readonly themeTint?: string;
  /** Theme shade (hex, e.g. "BF"). */
  readonly themeShade?: string;
}

/** Run properties (character-level formatting). */
export interface RunProperties {
  /** Font specification. */
  readonly font?: FontSpec | string;
  /** Font size in half-points. val=24 -> 12pt. */
  readonly size?: HalfPoint;
  /** Complex script font size in half-points. */
  readonly sizeCs?: HalfPoint;
  /** Bold. */
  readonly bold?: boolean;
  /** Bold for complex scripts. */
  readonly boldCs?: boolean;
  /** Italic. */
  readonly italic?: boolean;
  /** Italic for complex scripts. */
  readonly italicCs?: boolean;
  /** Underline style, boolean shorthand, or full spec with color. */
  readonly underline?: UnderlineStyle | boolean | UnderlineSpec;
  /** Strikethrough. */
  readonly strike?: boolean;
  /** Double strikethrough. */
  readonly doubleStrike?: boolean;
  /** All caps. */
  readonly caps?: boolean;
  /** Small caps. */
  readonly smallCaps?: boolean;
  /** Text color (hex RGB, "auto", or full spec with theme support). */
  readonly color?: HexColor | ColorSpec;
  /** Highlight color. */
  readonly highlight?: HighlightColor;
  /** Background shading. */
  readonly shading?: Shading;
  /** Superscript / subscript. */
  readonly vertAlign?: VerticalAlign;
  /** Character spacing in twips. */
  readonly spacing?: Twips;
  /** Character style ID. */
  readonly style?: string;
  /** Language. */
  readonly language?: {
    readonly val?: string;
    readonly eastAsia?: string;
    readonly bidi?: string;
  };
  /** Vanish (hidden text). */
  readonly vanish?: boolean;

  // --- Newly added properties ---

  /** Emboss effect. */
  readonly emboss?: boolean;
  /** Imprint (engrave) effect. */
  readonly imprint?: boolean;
  /** Do not check spelling/grammar. */
  readonly noProof?: boolean;
  /** Snap to document grid. */
  readonly snapToGrid?: boolean;
  /** Special vanish (hidden but still affects layout). */
  readonly specVanish?: boolean;
  /** Kerning threshold in half-points. Characters at or above this size will be kerned. */
  readonly kern?: HalfPoint;
  /** Vertical position offset in half-points (distinct from vertAlign). */
  readonly position?: HalfPoint;
  /** Character width scaling percentage (e.g. 100 = normal, 200 = double width). */
  readonly scale?: number;
  /** Right-to-left text. */
  readonly rightToLeft?: boolean;
  /** Complex script toggle (use CS font / bold / italic). */
  readonly complexScript?: boolean;
  /** Math mode flag. */
  readonly math?: boolean;
  /** Text effect / animation. */
  readonly effect?: TextEffect;
  /** Emphasis mark (East Asian typography). */
  readonly emphasisMark?: EmphasisMarkType;
  /** East Asian layout overrides (combine, vertical, etc.). ECMA-376 §17.3.2.10. */
  readonly eastAsianLayout?: EastAsianLayoutSpec;
  /** Character border (border around individual run). */
  readonly border?: Border;
  /** Fit text: force text to fit a specific width (in twips). */
  readonly fitText?: {
    /** Width in twips. */
    readonly val: Twips;
    /** Grouping ID for consecutive fitText runs. */
    readonly id?: number;
  };
  /** Outline (display only character outlines). */
  readonly outline?: boolean;
  /** Shadow effect on text. */
  readonly shadow?: boolean;
  /** Web hidden (hidden in web layout view). */
  readonly webHidden?: boolean;
  /** Run property change revision (track changes). */
  readonly propertyChange?: RunPropertyChange;
}

// =============================================================================
// Run Content
// =============================================================================

/** Text content within a run. */
export interface TextContent {
  readonly type: "text";
  /** Text string. */
  readonly text: string;
}

/** Line break. */
export interface BreakContent {
  readonly type: "break";
  /** Break type. Default: line break (no breakType = line break). */
  readonly breakType?: "page" | "column" | "textWrapping";
}

/** Tab character. */
export interface TabContent {
  readonly type: "tab";
}

/** Ruby alignment. */
export type RubyAlign =
  | "center"
  | "distributeLetter"
  | "distributeSpace"
  | "left"
  | "right"
  | "rightVertical";

/** Ruby (phonetic guide) properties. */
export interface RubyProperties {
  /** Alignment of ruby text over base. */
  readonly align?: RubyAlign;
  /** Ruby text font size (half-points). */
  readonly fontSize?: HalfPoint;
  /** Raise of ruby text (half-points above base). */
  readonly raise?: HalfPoint;
  /** Base text font size (half-points). */
  readonly baseFontSize?: HalfPoint;
  /** Language ID (e.g. "ja-JP", "zh-CN"). */
  readonly language?: string;
}

/** Ruby (phonetic guide) content — e.g. Japanese furigana or Chinese pinyin. */
export interface RubyContent {
  readonly type: "ruby";
  /** Ruby properties (styling). */
  readonly properties?: RubyProperties;
  /** Ruby text (the small phonetic text, shown above). */
  readonly rubyText: readonly Run[];
  /** Base text (the main text being annotated). */
  readonly baseText: readonly Run[];
}

/** Positional tab character (w:ptab). */
export interface PositionalTabContent {
  readonly type: "ptab";
  /** Alignment of content at tab. */
  readonly alignment: PositionalTabAlignment;
  /** Base of the positioning. */
  readonly relativeTo: PositionalTabRelativeTo;
  /** Leader character. */
  readonly leader?: PositionalTabLeader;
}

/** Symbol character. */
export interface SymbolContent {
  readonly type: "symbol";
  readonly font: string;
  readonly char: string;
}

/** Footnote reference within a run. */
export interface FootnoteRefContent {
  readonly type: "footnoteRef";
  readonly id: number;
  /** If true, the reference is followed by a custom mark character. */
  readonly customMarkFollows?: boolean;
}

/** Endnote reference within a run. */
export interface EndnoteRefContent {
  readonly type: "endnoteRef";
  readonly id: number;
  /** If true, the reference is followed by a custom mark character. */
  readonly customMarkFollows?: boolean;
}

/** Field code (PAGE, NUMPAGES, DATE, TOC, etc.). */
export interface FieldContent {
  readonly type: "field";
  /** Field instruction text (e.g. " PAGE ", " NUMPAGES "). */
  readonly instruction: string;
  /** Cached display value. */
  readonly cachedValue?: string;
  /** Legacy form field data (from w:ffData). */
  readonly formField?: FormField;
}

/** Inline image within a run. */
export interface InlineImageContent {
  readonly type: "image";
  /** Relationship ID for the image. */
  readonly rId: string;
  /** Width in EMU. */
  readonly width: Emu;
  /** Height in EMU. */
  readonly height: Emu;
  /** Alternative text for accessibility. */
  readonly altText?: string;
  /** Image name. */
  readonly name?: string;
  /** Unique drawing ID. */
  readonly drawingId?: number;
  /** Rotation in 60,000ths of a degree. */
  readonly rotation?: number;
  /** Flip horizontal. */
  readonly flipHorizontal?: boolean;
  /** Flip vertical. */
  readonly flipVertical?: boolean;
  /** Outline line properties. */
  readonly outline?: {
    readonly width?: Emu;
    readonly color?: HexColor;
  };
  /** SVG image relationship ID (for SVG images with raster fallback). */
  readonly svgRId?: string;
  /** Source rectangle for cropping (fractions: 100000 = 100%). */
  readonly srcRect?: {
    readonly l?: number;
    readonly t?: number;
    readonly r?: number;
    readonly b?: number;
  };
}

/** Carriage return element. */
export interface CarriageReturnContent {
  readonly type: "carriageReturn";
}

/** No-break hyphen. */
export interface NoBreakHyphenContent {
  readonly type: "noBreakHyphen";
}

/** Soft hyphen. */
export interface SoftHyphenContent {
  readonly type: "softHyphen";
}

/** Last rendered page break (read-only, informational). */
export interface LastRenderedPageBreakContent {
  readonly type: "lastRenderedPageBreak";
}

/** Annotation reference (for comments). */
export interface AnnotationReferenceContent {
  readonly type: "annotationReference";
  readonly id: number;
}

/** Page number constants for convenience fields. */
export type PageNumberType = "current" | "totalPages" | "totalPagesInSection" | "currentSection";

/** Date/time field. */
export interface DateFieldContent {
  readonly type: "dateField";
  /** Date format string (e.g. "yyyy-MM-dd"). */
  readonly format?: string;
  /** Language tag. */
  readonly language?: string;
  /** Pre-computed cached value to display. If omitted, the field shows empty until updated by Word. */
  readonly cachedValue?: string;
}

/** Content that can appear inside a run. */
export type RunContent =
  | TextContent
  | BreakContent
  | TabContent
  | PositionalTabContent
  | RubyContent
  | SymbolContent
  | FootnoteRefContent
  | EndnoteRefContent
  | FieldContent
  | InlineImageContent
  | CarriageReturnContent
  | NoBreakHyphenContent
  | SoftHyphenContent
  | LastRenderedPageBreakContent
  | AnnotationReferenceContent
  | DateFieldContent
  | OpaqueRunContent;

/**
 * Opaque run content: preserves unknown XML elements found inside a run
 * for round-trip fidelity. The writer will emit the raw XML verbatim.
 */
export interface OpaqueRunContent {
  readonly type: "opaqueRun";
  /** Raw XML string of the unrecognized run child element. */
  readonly rawXml: string;
}

// =============================================================================
// Run
// =============================================================================

/** A run of text with uniform formatting. */
export interface Run {
  /** Run properties (formatting). */
  readonly properties?: RunProperties;
  /** Content items within this run. */
  readonly content: readonly RunContent[];
}

// =============================================================================
// Track Changes / Revisions
// =============================================================================

/** Base revision information. */
export interface RevisionInfo {
  /** Author name. */
  readonly author: string;
  /** Date (ISO 8601 string). */
  readonly date?: string;
  /** Revision ID. */
  readonly id: number;
}

/** An inserted text run (track changes). */
export interface InsertedRun {
  readonly type: "insertedRun";
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** The run data. */
  readonly run: Run;
}

/** A deleted text run (track changes). */
export interface DeletedRun {
  readonly type: "deletedRun";
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** The run data (contains the deleted text for display). */
  readonly run: Run;
}

/** A moved-from run (track changes — source of a move). */
export interface MovedFromRun {
  readonly type: "movedFromRun";
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** The run data. */
  readonly run: Run;
}

/** A moved-to run (track changes — destination of a move). */
export interface MovedToRun {
  readonly type: "movedToRun";
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** The run data. */
  readonly run: Run;
}

/** Move range marker (w:moveFromRangeStart/End, w:moveToRangeStart/End). */
export interface MoveRangeMarker {
  readonly type: "moveFromRangeStart" | "moveFromRangeEnd" | "moveToRangeStart" | "moveToRangeEnd";
  /** Revision ID. */
  readonly id: number;
  /** Author (required for start, optional for end). */
  readonly author?: string;
  /** Date. */
  readonly date?: string;
  /** Name of the move (for pairing start/end). */
  readonly name?: string;
}

/** Custom XML tracking range marker. */
export interface CustomXmlTrackingMarker {
  readonly type:
    | "customXmlInsRangeStart"
    | "customXmlInsRangeEnd"
    | "customXmlDelRangeStart"
    | "customXmlDelRangeEnd"
    | "customXmlMoveFromRangeStart"
    | "customXmlMoveFromRangeEnd"
    | "customXmlMoveToRangeStart"
    | "customXmlMoveToRangeEnd";
  readonly id: number;
  readonly author?: string;
  readonly date?: string;
}

/** Paragraph property change revision. */
export interface ParagraphPropertyChange {
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** Previous paragraph properties. */
  readonly previousProperties?: ParagraphProperties;
}

/** Run property change revision. */
export interface RunPropertyChange {
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** Previous run properties. */
  readonly previousProperties?: RunProperties;
}

/** Section property change revision. */
export interface SectionPropertyChange {
  /** Revision metadata. */
  readonly revision: RevisionInfo;
  /** Previous section properties. */
  readonly previousProperties?: SectionProperties;
}

/** Table property change revision (w:tblPrChange). */
export interface TablePropertyChange {
  readonly revision: RevisionInfo;
  readonly previousProperties?: TableProperties;
}

/** Table row property change revision (w:trPrChange). */
export interface TableRowPropertyChange {
  readonly revision: RevisionInfo;
  readonly previousProperties?: TableRowProperties;
}

/** Table cell property change revision (w:tcPrChange). */
export interface TableCellPropertyChange {
  readonly revision: RevisionInfo;
  readonly previousProperties?: TableCellProperties;
}

/** Table grid cell revision (w:cellIns / w:cellDel / w:cellMerge). */
export interface CellMergeRevision {
  /** Merge type. */
  readonly vMerge: "cont" | "rest";
  readonly revision: RevisionInfo;
}

/** Table row insertion/deletion revision. */
export interface TableRowRevision {
  /** Revision metadata. */
  readonly revision: RevisionInfo;
}

// =============================================================================
// Comments
// =============================================================================

/** A comment definition. */
export interface CommentDef {
  /** Unique comment ID. */
  readonly id: number;
  /** Author name. */
  readonly author: string;
  /** Date (ISO 8601 string). */
  readonly date?: string;
  /** Author initials. */
  readonly initials?: string;
  /** Comment content (paragraphs). */
  readonly content: readonly Paragraph[];
  /** Whether the comment has been marked resolved/done (w15:done). */
  readonly done?: boolean;
  /** Parent comment paraId (for reply threads; w15:paraIdParent). */
  readonly parentId?: string;
}

/** Comment range start marker. */
export interface CommentRangeStart {
  readonly type: "commentRangeStart";
  readonly id: number;
}

/** Comment range end marker. */
export interface CommentRangeEnd {
  readonly type: "commentRangeEnd";
  readonly id: number;
}

/** Comment reference (in a run, links to a comment). */
export interface CommentReference {
  readonly type: "commentReference";
  readonly id: number;
}

// =============================================================================
// Paragraph Properties
// =============================================================================

/** Horizontal alignment (ST_Jc). */
export type Alignment =
  | "start"
  | "center"
  | "end"
  | "both"
  | "mediumKashida"
  | "distribute"
  | "numTab"
  | "highKashida"
  | "lowKashida"
  | "thaiDistribute"
  | "left"
  | "right";

/** Line spacing rule. */
export type LineSpacingRule = "auto" | "exact" | "atLeast";

/** Line spacing. */
export interface LineSpacing {
  /** Spacing value. For "auto": 240ths of a line (240=single, 360=1.5, 480=double). For "exact"/"atLeast": twips. */
  readonly line?: number;
  /** Spacing rule. Default: "auto". */
  readonly lineRule?: LineSpacingRule;
  /** Spacing before paragraph in twips. */
  readonly before?: Twips;
  /** Spacing after paragraph in twips. */
  readonly after?: Twips;
  /** Automatic spacing before. */
  readonly beforeAutoSpacing?: boolean;
  /** Automatic spacing after. */
  readonly afterAutoSpacing?: boolean;
}

/** Paragraph indentation in twips. */
export interface Indentation {
  readonly left?: Twips;
  readonly right?: Twips;
  /** Hanging indent (mutually exclusive with firstLine). */
  readonly hanging?: Twips;
  /** First line indent (mutually exclusive with hanging). */
  readonly firstLine?: Twips;
  /** Start indent (logical direction, for BiDi). */
  readonly start?: Twips;
  /** End indent (logical direction, for BiDi). */
  readonly end?: Twips;
}

/** Tab stop alignment (ST_TabJc). */
export type TabStopType =
  | "left"
  | "center"
  | "right"
  | "decimal"
  | "bar"
  | "clear"
  | "start"
  | "end"
  | "num";

/** Tab stop leader character. */
export type TabStopLeader = "dot" | "hyphen" | "underscore" | "none" | "heavy" | "middleDot";

/** Tab stop definition. */
export interface TabStop {
  readonly type: TabStopType;
  /** Position in twips. */
  readonly position: Twips;
  readonly leader?: TabStopLeader;
}

/** Paragraph borders. */
export interface ParagraphBorders {
  readonly top?: Border;
  readonly bottom?: Border;
  readonly left?: Border;
  readonly right?: Border;
  readonly between?: Border;
  readonly bar?: Border;
}

/** Numbering reference for lists. */
export interface NumberingRef {
  /** Numbering instance ID. */
  readonly numId: number;
  /** Indentation level (0-8). */
  readonly level: number;
}

/** Frame anchor type. */
export type FrameAnchorType = "margin" | "page" | "text";

/** Drop cap type. */
export type DropCapType = "drop" | "margin" | "none";

/** Paragraph frame properties. */
export interface ParagraphFrame {
  /** Drop cap type. */
  readonly dropCap?: DropCapType;
  /** Number of lines for drop cap. */
  readonly lines?: number;
  /** Width in twips. */
  readonly width?: Twips;
  /** Height in twips. */
  readonly height?: Twips;
  /** Horizontal space (margin from text) in twips. */
  readonly hSpace?: Twips;
  /** Vertical space in twips. */
  readonly vSpace?: Twips;
  /** Wrap text around frame. */
  readonly wrap?: "auto" | "around" | "none" | "notBeside" | "through" | "tight";
  /** Horizontal anchor. */
  readonly hAnchor?: FrameAnchorType;
  /** Vertical anchor. */
  readonly vAnchor?: FrameAnchorType;
  /** X position in twips. */
  readonly x?: Twips;
  /** Horizontal alignment. */
  readonly xAlign?: "left" | "center" | "right" | "inside" | "outside";
  /** Y position in twips. */
  readonly y?: Twips;
  /** Vertical alignment. */
  readonly yAlign?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
}

/** Positional tab alignment. */
export type PositionalTabAlignment = "center" | "left" | "right";

/** Positional tab relative to. */
export type PositionalTabRelativeTo = "indent" | "margin";

/** Positional tab leader. */
export type PositionalTabLeader = "dot" | "hyphen" | "middleDot" | "none" | "underscore";

/** Paragraph properties. */
export interface ParagraphProperties {
  /** Paragraph style ID. */
  readonly style?: string;
  /** Alignment. */
  readonly alignment?: Alignment;
  /** Indentation. */
  readonly indent?: Indentation;
  /** Spacing. */
  readonly spacing?: LineSpacing;
  /** Numbering/list reference. */
  readonly numbering?: NumberingRef;
  /** Paragraph borders. */
  readonly borders?: ParagraphBorders;
  /** Background shading. */
  readonly shading?: Shading;
  /** Tab stops. */
  readonly tabs?: readonly TabStop[];
  /** Keep with next paragraph on same page. */
  readonly keepNext?: boolean;
  /** Keep all lines on same page. */
  readonly keepLines?: boolean;
  /** Page break before paragraph. */
  readonly pageBreakBefore?: boolean;
  /** Widow/orphan control. */
  readonly widowControl?: boolean;
  /** Bidirectional text. */
  readonly bidi?: boolean;
  /** Section properties (for section breaks within paragraphs). */
  readonly sectionProperties?: SectionProperties;
  /** Run properties applied to the paragraph mark. */
  readonly markRunProperties?: RunProperties;

  // --- Newly added properties ---

  /** Contextual spacing (ignore before/after for same-style paragraphs). */
  readonly contextualSpacing?: boolean;
  /** Thematic break (horizontal rule below paragraph). */
  readonly thematicBreak?: boolean;
  /** Suppress line numbers for this paragraph. */
  readonly suppressLineNumbers?: boolean;
  /** Allow automatic word wrap. */
  readonly wordWrap?: boolean;
  /** Allow punctuation to overflow margins (East Asian). */
  readonly overflowPunctuation?: boolean;
  /** Auto-space East Asian and Latin text. */
  readonly autoSpaceEastAsianText?: boolean;
  /** Auto-space between East Asian text and numbers (w:autoSpaceDN). */
  readonly autoSpaceEastAsianDigit?: boolean;
  /** Disable kinsoku (East Asian line-break rules). When false, kinsoku is enabled. */
  readonly kinsoku?: boolean;
  /** Top-line punctuation compression (East Asian). */
  readonly topLinePunctuation?: boolean;
  /** Outline level (0-9, used for TOC). */
  readonly outlineLevel?: number;
  /** Frame properties (paragraph positioning). */
  readonly frame?: ParagraphFrame;
  /** Paragraph property change revision (track changes). */
  readonly propertyChange?: ParagraphPropertyChange;
  /** Paragraph insertion revision (the paragraph mark itself was inserted). */
  readonly paragraphInsertion?: RevisionInfo;
  /** Paragraph deletion revision (the paragraph mark itself was deleted). */
  readonly paragraphDeletion?: RevisionInfo;
  /** Suppress auto-hyphens. */
  readonly suppressAutoHyphens?: boolean;
  /** Conditional formatting style mask (w:cnfStyle). */
  readonly cnfStyle?: string;
  /** Mirror indents for odd/even pages. */
  readonly mirrorIndents?: boolean;
  /** Text alignment within line (for East Asian vertical text). */
  readonly textAlignment?: "auto" | "baseline" | "bottom" | "center" | "top";
  /** Snap to grid. */
  readonly snapToGrid?: boolean;
  /** Text direction for paragraph. */
  readonly textDirection?: TextDirection;
}

// =============================================================================
// Block-Level Content
// =============================================================================

/** A paragraph. */
export interface Paragraph {
  readonly type: "paragraph";
  /** Paragraph properties. */
  readonly properties?: ParagraphProperties;
  /** Runs and hyperlinks within this paragraph. */
  readonly children: readonly ParagraphChild[];
  /** Paragraph ID (w14:paraId) for commenting/collaboration features. */
  readonly paraId?: string;
  /** Text ID (w14:textId) for commenting/collaboration features. */
  readonly textId?: string;
}

/** A hyperlink within a paragraph. */
export interface Hyperlink {
  readonly type: "hyperlink";
  /** Relationship ID for external hyperlinks. */
  readonly rId?: string;
  /** Resolved URL for external hyperlinks (populated by reader, used by builder). */
  readonly url?: string;
  /** Anchor name for internal bookmark links. */
  readonly anchor?: string;
  /** Tooltip text displayed on hover. */
  readonly tooltip?: string;
  /** Whether the hyperlink has been visited (w:history). */
  readonly history?: boolean;
  /** Target frame name for web views (w:tgtFrame). */
  readonly tgtFrame?: string;
  /** Document location bookmark/fragment ID (w:docLocation). */
  readonly docLocation?: string;
  /** Runs within the hyperlink. */
  readonly children: readonly Run[];
}

/** Bookmark start marker. */
export interface BookmarkStart {
  readonly type: "bookmarkStart";
  readonly id: number;
  readonly name: string;
  /** Column first (for column bookmarks in tables). */
  readonly colFirst?: number;
  /** Column last (for column bookmarks in tables). */
  readonly colLast?: number;
  /** Displaced by custom XML. */
  readonly displacedByCustomXml?: "next" | "prev";
}

/** Bookmark end marker. */
export interface BookmarkEnd {
  readonly type: "bookmarkEnd";
  readonly id: number;
}

/** Items that can appear as children of a paragraph. */
export type ParagraphChild =
  | Run
  | Hyperlink
  | BookmarkStart
  | BookmarkEnd
  | CommentRangeStart
  | CommentRangeEnd
  | CommentReference
  | InsertedRun
  | DeletedRun
  | MovedFromRun
  | MovedToRun
  | MoveRangeMarker
  | CustomXmlTrackingMarker
  | OpaqueParagraphChild;

/**
 * Opaque paragraph child: preserves unknown XML elements found at paragraph level
 * for round-trip fidelity. The writer will emit the raw XML verbatim.
 */
export interface OpaqueParagraphChild {
  readonly type: "opaqueParagraphChild";
  /** Raw XML string of the unrecognized paragraph child element. */
  readonly rawXml: string;
}

// =============================================================================
// Table
// =============================================================================

/** Table width specification. */
export interface TableWidth {
  /** Width value. */
  readonly value: number;
  /** Width type: "auto", "dxa" (twips), "pct" (fiftieths of a percent, 5000=100%). */
  readonly type: "auto" | "dxa" | "pct" | "nil";
}

/** Table borders. */
export interface TableBorders {
  readonly top?: Border;
  readonly left?: Border;
  readonly bottom?: Border;
  readonly right?: Border;
  readonly insideH?: Border;
  readonly insideV?: Border;
  /** Start border (logical direction). */
  readonly start?: Border;
  /** End border (logical direction). */
  readonly end?: Border;
  /** Top-left to bottom-right diagonal border (cells only). */
  readonly tl2br?: Border;
  /** Top-right to bottom-left diagonal border (cells only). */
  readonly tr2bl?: Border;
}

/** Table cell margins. */
export interface TableCellMargins {
  readonly top?: TableWidth;
  readonly left?: TableWidth;
  readonly bottom?: TableWidth;
  readonly right?: TableWidth;
  /** Start margin (logical). */
  readonly start?: TableWidth;
  /** End margin (logical). */
  readonly end?: TableWidth;
}

/** Table layout algorithm. */
export type TableLayout = "fixed" | "autofit";

/** Table look (conditional formatting flags). */
export interface TableLook {
  /** Apply first row conditional formatting. */
  readonly firstRow?: boolean;
  /** Apply last row conditional formatting. */
  readonly lastRow?: boolean;
  /** Apply first column conditional formatting. */
  readonly firstColumn?: boolean;
  /** Apply last column conditional formatting. */
  readonly lastColumn?: boolean;
  /** Do not apply horizontal banding. */
  readonly noHBand?: boolean;
  /** Do not apply vertical banding. */
  readonly noVBand?: boolean;
}

/** Table floating properties. */
export interface TableFloat {
  /** Horizontal anchor: "margin" | "page" | "text". */
  readonly horizontalAnchor?: "margin" | "page" | "text";
  /** Vertical anchor: "margin" | "page" | "text". */
  readonly verticalAnchor?: "margin" | "page" | "text";
  /** Horizontal absolute position in twips. */
  readonly absoluteHorizontalPosition?: Twips;
  /** Vertical absolute position in twips. */
  readonly absoluteVerticalPosition?: Twips;
  /** Relative horizontal position. */
  readonly relativeHorizontalPosition?: "center" | "inside" | "left" | "outside" | "right";
  /** Relative vertical position. */
  readonly relativeVerticalPosition?: "center" | "inside" | "bottom" | "outside" | "top" | "inline";
  /** Top margin from text in twips. */
  readonly topFromText?: Twips;
  /** Bottom margin from text in twips. */
  readonly bottomFromText?: Twips;
  /** Left margin from text in twips. */
  readonly leftFromText?: Twips;
  /** Right margin from text in twips. */
  readonly rightFromText?: Twips;
  /** Allow overlap with other floating tables. */
  readonly overlap?: "never" | "overlap";
}

/** Table properties. */
export interface TableProperties {
  /** Table style ID. */
  readonly style?: string;
  /** Table width. */
  readonly width?: TableWidth;
  /** Table borders. */
  readonly borders?: TableBorders;
  /** Table layout. */
  readonly layout?: TableLayout;
  /** Cell margins (defaults for all cells). */
  readonly cellMargins?: TableCellMargins;
  /** Table alignment. */
  readonly alignment?: Alignment;
  /** Table indent in twips. */
  readonly indent?: Twips;
  /** Table look (conditional formatting). */
  readonly look?: TableLook;
  /** Floating table properties. */
  readonly float?: TableFloat;
  /** Cell spacing in twips. */
  readonly cellSpacing?: TableWidth;
  /** Visual right-to-left. */
  readonly visuallyRightToLeft?: boolean;
  /** Background shading for entire table. */
  readonly shading?: Shading;
  /** Table caption (accessibility). */
  readonly caption?: string;
  /** Table description (accessibility). */
  readonly description?: string;
  /** Table property change revision (track changes). */
  readonly propertyChange?: TablePropertyChange;
}

/** Table row height rule. */
export type RowHeightRule = "auto" | "atLeast" | "exact";

/** Table row properties. */
export interface TableRowProperties {
  /** Row height. */
  readonly height?: { readonly value: Twips; readonly rule?: RowHeightRule };
  /** Repeat as header row across pages. */
  readonly tableHeader?: boolean;
  /** Don't split row across pages. */
  readonly cantSplit?: boolean;
  /** Cell spacing for this row. */
  readonly cellSpacing?: TableWidth;
  /** Hidden row (track changes). */
  readonly hidden?: boolean;
  /** Row insertion revision (track changes). */
  readonly inserted?: TableRowRevision;
  /** Row deletion revision (track changes). */
  readonly deleted?: TableRowRevision;
  /** Number of grid units before the first cell (w:gridBefore). */
  readonly gridBefore?: number;
  /** Number of grid units after the last cell (w:gridAfter). */
  readonly gridAfter?: number;
  /** Width of the grid before (w:wBefore). */
  readonly widthBefore?: TableWidth;
  /** Width of the grid after (w:wAfter). */
  readonly widthAfter?: TableWidth;
  /** Conditional formatting style mask (w:cnfStyle). */
  readonly cnfStyle?: string;
  /** Table row property change revision (track changes). */
  readonly propertyChange?: TableRowPropertyChange;
  /** Row-level table property exceptions (w:tblPrEx) — override of table props. */
  readonly tblPrEx?: TableProperties;
}

/** Vertical alignment in a cell. */
export type VerticalCellAlign = "top" | "center" | "bottom";

/** Text direction in a cell. */
export type TextDirection = "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV" | "tbLrV";

/** Vertical merge type. */
export type VerticalMerge = "restart" | "continue";

/** Table cell properties. */
export interface TableCellProperties {
  /** Cell width. */
  readonly width?: TableWidth;
  /** Horizontal span (gridSpan). */
  readonly gridSpan?: number;
  /** Vertical merge. "restart" starts a merge, "continue" continues it. */
  readonly verticalMerge?: VerticalMerge;
  /** Cell borders. */
  readonly borders?: TableBorders;
  /** Cell background shading. */
  readonly shading?: Shading;
  /** Vertical alignment. */
  readonly verticalAlign?: VerticalCellAlign;
  /** Text direction. */
  readonly textDirection?: TextDirection;
  /** No text wrapping. */
  readonly noWrap?: boolean;
  /** Cell margins (overrides table defaults). */
  readonly margins?: TableCellMargins;
  /** Convenience rowSpan. Used by builder to auto-generate vMerge. */
  readonly rowSpan?: number;
  /** Conditional formatting style mask (w:cnfStyle). */
  readonly cnfStyle?: string;
  /** Hide cell end-of-cell marker (w:hideMark). */
  readonly hideMark?: boolean;
  /** Fits text in cell width (w:tcFitText). */
  readonly fitText?: boolean;
  /** Cell property change revision (track changes). */
  readonly propertyChange?: TableCellPropertyChange;
  /** Cell inserted revision (w:cellIns — cell was inserted). */
  readonly inserted?: TableRowRevision;
  /** Cell deleted revision (w:cellDel — cell was deleted). */
  readonly deleted?: TableRowRevision;
  /** Cell merge revision (w:cellMerge). */
  readonly cellMerge?: CellMergeRevision;
}

/** Table cell. */
export interface TableCell {
  /** Cell properties. */
  readonly properties?: TableCellProperties;
  /** Cell content (must contain at least one paragraph). */
  readonly content: readonly (Paragraph | Table)[];
}

/** Table row. */
export interface TableRow {
  /** Row properties. */
  readonly properties?: TableRowProperties;
  /** Cells in this row. */
  readonly cells: readonly TableCell[];
}

/** A table. */
export interface Table {
  readonly type: "table";
  /** Table properties. */
  readonly properties?: TableProperties;
  /** Column widths in twips (for tblGrid). */
  readonly columnWidths?: readonly Twips[];
  /** Rows. */
  readonly rows: readonly TableRow[];
}

// =============================================================================
// Floating Image
// =============================================================================

/** Horizontal position relative to. */
export type HorizontalPositionRelative =
  | "character"
  | "column"
  | "insideMargin"
  | "leftMargin"
  | "margin"
  | "outsideMargin"
  | "page"
  | "rightMargin";

/** Vertical position relative to. */
export type VerticalPositionRelative =
  | "insideMargin"
  | "line"
  | "margin"
  | "outsideMargin"
  | "page"
  | "paragraph"
  | "topMargin"
  | "bottomMargin";

/** Text wrapping style. */
export type WrapStyle = "square" | "tight" | "through" | "topAndBottom" | "none";

/** Wrap text side. */
export type WrapTextSide = "bothSides" | "left" | "right" | "largest";

/** Wrap distance margins. */
export interface WrapMargins {
  readonly top?: Emu;
  readonly bottom?: Emu;
  readonly left?: Emu;
  readonly right?: Emu;
}

/** Floating image (anchor). */
export interface FloatingImage {
  readonly type: "floatingImage";
  /** Relationship ID for the image. */
  readonly rId: string;
  /** Width in EMU. */
  readonly width: Emu;
  /** Height in EMU. */
  readonly height: Emu;
  /** Alternative text. */
  readonly altText?: string;
  /** Image name. */
  readonly name?: string;
  /** Horizontal position. */
  readonly horizontalPosition?: {
    readonly relativeTo?: HorizontalPositionRelative;
    readonly offset?: Emu;
    readonly align?: "left" | "center" | "right" | "inside" | "outside";
  };
  /** Vertical position. */
  readonly verticalPosition?: {
    readonly relativeTo?: VerticalPositionRelative;
    readonly offset?: Emu;
    readonly align?: "top" | "center" | "bottom" | "inside" | "outside";
  };
  /** Text wrapping. */
  readonly wrap?: {
    readonly style: WrapStyle;
    readonly side?: WrapTextSide;
    /** Wrap distance margins. */
    readonly margins?: WrapMargins;
  };
  /** Z-order relative height. */
  readonly relativeHeight?: number;
  /** Place behind document text. */
  readonly behindDoc?: boolean;
  /** Unique drawing ID. */
  readonly drawingId?: number;
  /** Rotation in 60,000ths of a degree. */
  readonly rotation?: number;
  /** Flip horizontal. */
  readonly flipHorizontal?: boolean;
  /** Flip vertical. */
  readonly flipVertical?: boolean;
  /** Lock anchor position. */
  readonly lockAnchor?: boolean;
  /** Layout in cell (default: true). */
  readonly layoutInCell?: boolean;
  /** Allow overlap with other floating content (default: true). */
  readonly allowOverlap?: boolean;
  /** Simple positioning with absolute x/y (alternative to relative anchor positioning). */
  readonly simplePos?: {
    readonly x: Emu;
    readonly y: Emu;
  };
  /** Distance from top/bottom/left/right of surrounding text (EMU). */
  readonly distT?: Emu;
  readonly distB?: Emu;
  readonly distL?: Emu;
  readonly distR?: Emu;
  /** Outline / border. */
  readonly outline?: {
    readonly width?: Emu;
    readonly color?: HexColor;
  };
  /** SVG image relationship ID (for SVG images with raster fallback). */
  readonly svgRId?: string;
  /** Source rectangle for cropping (fractions: 100000 = 100%). */
  readonly srcRect?: {
    readonly l?: number;
    readonly t?: number;
    readonly r?: number;
    readonly b?: number;
  };
}

// =============================================================================
// Table of Contents (TOC)
// =============================================================================

/** Table of Contents properties. */
export interface TableOfContents {
  readonly type: "tableOfContents";
  /** Heading style range (e.g. "1-3" for Heading 1 through 3). */
  readonly headingStyleRange?: string;
  /** Whether entries are hyperlinks. */
  readonly hyperlink?: boolean;
  /** Styles with levels (e.g. [{ styleName: "MyStyle", level: 1 }]). */
  readonly stylesWithLevels?: readonly {
    readonly styleName: string;
    readonly level: number;
  }[];
  /** Caption label (for table of figures). */
  readonly captionLabel?: string;
  /** Sequence field identifier (for table of figures). */
  readonly sequenceFieldIdentifier?: string;
  /** Tab stop leader for page numbers. */
  readonly leader?: TabStopLeader;
  /** Cached/fallback paragraphs to display when not updated by Word. */
  readonly cachedParagraphs?: readonly Paragraph[];
}

// =============================================================================
// Math / Equations (OMML)
// =============================================================================

/** Math run (text in math context). */
export interface MathRun {
  readonly type: "mathRun";
  /** Math text content. */
  readonly text: string;
  /** Math run properties (italic, font, etc.). */
  readonly properties?: {
    readonly italic?: boolean;
    readonly bold?: boolean;
    readonly font?: string;
  };
}

/** Math fraction. */
export interface MathFraction {
  readonly type: "mathFraction";
  /** Fraction type. */
  readonly fractionType?: "bar" | "skw" | "lin" | "noBar";
  /** Numerator. */
  readonly numerator: readonly MathContent[];
  /** Denominator. */
  readonly denominator: readonly MathContent[];
}

/** Math superscript. */
export interface MathSuperScript {
  readonly type: "mathSuperScript";
  readonly base: readonly MathContent[];
  readonly superScript: readonly MathContent[];
}

/** Math subscript. */
export interface MathSubScript {
  readonly type: "mathSubScript";
  readonly base: readonly MathContent[];
  readonly subScript: readonly MathContent[];
}

/** Math sub-superscript. */
export interface MathSubSuperScript {
  readonly type: "mathSubSuperScript";
  readonly base: readonly MathContent[];
  readonly subScript: readonly MathContent[];
  readonly superScript: readonly MathContent[];
}

/** Math pre-sub-superscript (m:sPre) — subscript and superscript attached before the base. */
export interface MathPreSubSuperScript {
  readonly type: "mathPreSubSuperScript";
  readonly base: readonly MathContent[];
  readonly preSubScript: readonly MathContent[];
  readonly preSuperScript: readonly MathContent[];
}

/** Math phantom (m:phant) — rendered invisibly but occupies space. */
export interface MathPhantom {
  readonly type: "mathPhantom";
  readonly content: readonly MathContent[];
  /** Phantom shows up in the layout (default: false). */
  readonly show?: boolean;
  /** Zero width (default: false). */
  readonly zeroWidth?: boolean;
  /** Zero ascent (default: false). */
  readonly zeroAscent?: boolean;
  /** Zero descent (default: false). */
  readonly zeroDescent?: boolean;
  /** Transparent (default: false). */
  readonly transparent?: boolean;
}

/** Math group character (m:groupChr) — a character grouping (e.g. horizontal brace). */
export interface MathGroupChar {
  readonly type: "mathGroupChar";
  readonly base: readonly MathContent[];
  /** The group character (e.g. "{" or "⏞"). */
  readonly char?: string;
  /** Position: top or bottom. */
  readonly position?: "top" | "bottom";
  /** Vertical alignment. */
  readonly verticalAlign?: "top" | "center" | "bottom";
}

/** Math border box (m:borderBox) — a box around the expression with optional borders. */
export interface MathBorderBox {
  readonly type: "mathBorderBox";
  readonly content: readonly MathContent[];
  readonly hideTop?: boolean;
  readonly hideBottom?: boolean;
  readonly hideLeft?: boolean;
  readonly hideRight?: boolean;
  /** Strike diagonals (bltr = bottom-left to top-right; tlbr = top-left to bottom-right; h = horizontal; v = vertical). */
  readonly strikeBlTr?: boolean;
  readonly strikeTlBr?: boolean;
  readonly strikeH?: boolean;
  readonly strikeV?: boolean;
}

/** Math radical (square root / nth root). */
export interface MathRadical {
  readonly type: "mathRadical";
  /** Degree (for nth root). If omitted, square root. */
  readonly degree?: readonly MathContent[];
  /** Content under the radical. */
  readonly content: readonly MathContent[];
  /** Hide the degree. */
  readonly hideDegree?: boolean;
}

/** Math delimiter (parentheses, brackets, etc.). */
export interface MathDelimiter {
  readonly type: "mathDelimiter";
  /** Beginning character (default: "("). */
  readonly beginChar?: string;
  /** Ending character (default: ")"). */
  readonly endChar?: string;
  /** Separator character (default: "|"). */
  readonly separatorChar?: string;
  /** Content elements. */
  readonly content: readonly (readonly MathContent[])[];
}

/** Math summation / product / integral with limits. */
export interface MathNary {
  readonly type: "mathNary";
  /** The operator character (e.g. "\u2211" for sum, "\u220F" for product, "\u222B" for integral). */
  readonly char?: string;
  /** Lower limit (subscript). */
  readonly sub?: readonly MathContent[];
  /** Upper limit (superscript). */
  readonly sup?: readonly MathContent[];
  /** Content. */
  readonly content: readonly MathContent[];
  /** Place limits above/below (true) or as sub/superscript (false). */
  readonly limitsLocation?: "subSup" | "undOvr";
  /** Hide upper limit. */
  readonly supHide?: boolean;
  /** Hide lower limit. */
  readonly subHide?: boolean;
}

/** Math function (e.g. sin, cos). */
export interface MathFunction {
  readonly type: "mathFunction";
  /** Function name content. */
  readonly name: readonly MathContent[];
  /** Function argument. */
  readonly content: readonly MathContent[];
}

/** Math limit (upper or lower). */
export interface MathLimit {
  readonly type: "mathLimit";
  /** "upper" or "lower". */
  readonly limitType: "upper" | "lower";
  /** Base content. */
  readonly base: readonly MathContent[];
  /** Limit content. */
  readonly limit: readonly MathContent[];
}

/** Math matrix. */
export interface MathMatrix {
  readonly type: "mathMatrix";
  /** Matrix rows, each containing cells. */
  readonly rows: readonly (readonly (readonly MathContent[])[])[];
}

/** Math accent (hat, tilde, etc.). */
export interface MathAccent {
  readonly type: "mathAccent";
  /** Accent character (e.g. "\u0302" for hat). */
  readonly char?: string;
  /** Base content. */
  readonly content: readonly MathContent[];
}

/** Math bar (overline / underline). */
export interface MathBar {
  readonly type: "mathBar";
  /** Bar position. */
  readonly position: "top" | "bottom";
  /** Content. */
  readonly content: readonly MathContent[];
}

/** Math box (invisible grouping). */
export interface MathBox {
  readonly type: "mathBox";
  readonly content: readonly MathContent[];
}

/** Math equation array (system of equations). */
export interface MathEquationArray {
  readonly type: "mathEquationArray";
  /** Array of equation rows. */
  readonly rows: readonly (readonly MathContent[])[];
}

/** All math content types. */
export type MathContent =
  | MathRun
  | MathFraction
  | MathSuperScript
  | MathSubScript
  | MathSubSuperScript
  | MathPreSubSuperScript
  | MathRadical
  | MathDelimiter
  | MathNary
  | MathFunction
  | MathLimit
  | MathMatrix
  | MathAccent
  | MathBar
  | MathBox
  | MathPhantom
  | MathGroupChar
  | MathBorderBox
  | MathEquationArray;

/** A math block (inline or display). */
export interface MathBlock {
  readonly type: "math";
  /** Math content elements. */
  readonly content: readonly MathContent[];
}

// =============================================================================
// Text Box
// =============================================================================

/** Text box definition. */
export interface TextBox {
  readonly type: "textBox";
  /** Content paragraphs. */
  readonly content: readonly Paragraph[];
  /** Width in twips. */
  readonly width?: Twips;
  /** Height in twips. */
  readonly height?: Twips;
  /** Style string for VML (e.g. CSS-like position/size). */
  readonly style?: string;
  /** Stroke visible. */
  readonly stroke?: boolean;
  /** Stroke color. */
  readonly strokeColor?: HexColor;
  /** Fill visible. */
  readonly fill?: boolean;
  /** Fill color. */
  readonly fillColor?: HexColor;
}

// =============================================================================
// DrawingML Shape (wsp:)
// =============================================================================

/** Common preset shape types. */
export type PresetShapeType = string;

/** DrawingML shape (wsp:). */
export interface DrawingShape {
  readonly type: "drawingShape";
  /** Preset shape type (e.g. "rect", "roundRect", "ellipse"). */
  readonly shapeType: PresetShapeType;
  /** Width in EMU. */
  readonly width: Emu;
  /** Height in EMU. */
  readonly height: Emu;
  /** Fill color (hex). */
  readonly fillColor?: HexColor;
  /** No fill. */
  readonly noFill?: boolean;
  /** Outline/stroke color (hex). */
  readonly outlineColor?: HexColor;
  /** Outline width in EMU. */
  readonly outlineWidth?: Emu;
  /** No outline. */
  readonly noOutline?: boolean;
  /** Text content inside the shape. */
  readonly textContent?: readonly Paragraph[];
  /** Alternative text. */
  readonly altText?: string;
  /** Shape name. */
  readonly name?: string;
  /** Horizontal position. */
  readonly horizontalPosition?: {
    readonly relativeTo?: HorizontalPositionRelative;
    readonly offset?: Emu;
    readonly align?: "left" | "center" | "right";
  };
  /** Vertical position. */
  readonly verticalPosition?: {
    readonly relativeTo?: VerticalPositionRelative;
    readonly offset?: Emu;
    readonly align?: "top" | "center" | "bottom";
  };
  /** Text wrapping. */
  readonly wrap?: {
    readonly style: WrapStyle;
    readonly side?: WrapTextSide;
  };
  /** Behind document text. */
  readonly behindDoc?: boolean;
  /** Rotation in 60,000ths of a degree. */
  readonly rotation?: number;
  /** Raw XML string (for preserving unrecognized shape details on round-trip). */
  readonly rawXml?: string;
  /**
   * Advanced fill XML (gradient/pattern) — must be inserted in the spPr
   * "fill" slot (after prstGeom, before a:ln) for OOXML schema validity.
   * Internal: populated by createShape().
   */
  readonly _advancedFillXml?: string;
  /**
   * Advanced effects/3D XML (a:effectLst, a:scene3d, a:sp3d) — must be
   * inserted in the spPr after a:ln to satisfy the OOXML schema order.
   * Internal: populated by createShape().
   */
  readonly _advancedEffectsXml?: string;
}

// =============================================================================
// CheckBox
// =============================================================================

/** Checkbox state. */
export interface CheckBox {
  readonly type: "checkBox";
  /** Whether checked. */
  readonly checked?: boolean;
  /** Checked state character. */
  readonly checkedState?: { readonly value: string; readonly font?: string };
  /** Unchecked state character. */
  readonly uncheckedState?: { readonly value: string; readonly font?: string };
}

// =============================================================================
// Structured Document Tags (SDT) — content controls
// =============================================================================

/** SDT dropdown/combobox list item. */
export interface SdtListItem {
  /** Display text. */
  readonly displayText?: string;
  /** Value. */
  readonly value: string;
}

/** SDT date picker properties. */
export interface SdtDateProperties {
  /** Full date value (ISO 8601). */
  readonly fullDate?: string;
  /** Date format string (e.g. "M/d/yyyy"). */
  readonly dateFormat?: string;
  /** Language ID. */
  readonly lid?: string;
  /** Storage type for date mappings. */
  readonly storeMappedDataAs?: "dateTime" | "date" | "text";
}

/** Structured Document Tag properties. */
export interface SdtProperties {
  /** SDT unique ID. */
  readonly id?: number;
  /** Tag name. */
  readonly tag?: string;
  /** Alias/title. */
  readonly alias?: string;
  /** Lock content from editing. */
  readonly lockContent?: boolean;
  /** Lock SDT from deletion. */
  readonly lockSdt?: boolean;
  /** Dropdown list items. */
  readonly dropdownList?: readonly SdtListItem[];
  /** ComboBox list items. */
  readonly comboBox?: readonly SdtListItem[];
  /** Date picker properties. */
  readonly date?: SdtDateProperties;
  /** Plain text content (vs rich text). */
  readonly plainText?: boolean;
  /** Appearance (w15:appearance): how the SDT is visually delineated. */
  readonly appearance?: "boundingBox" | "tags" | "hidden";
  /**
   * Show placeholder text toggle (w:showingPlcHdr). When true, the control
   * shows placeholder text; cleared automatically on first edit.
   */
  readonly showingPlaceholder?: boolean;
  /** Placeholder text doc part reference. */
  readonly placeholder?: string;
  /** Data binding to Custom XML part. */
  readonly dataBinding?: SdtDataBinding;
  /** Rich text SDT marker. */
  readonly richText?: boolean;
  /** Picture SDT marker. */
  readonly picture?: boolean;
  /** Grouping SDT marker (groups content, prevents editing). */
  readonly group?: boolean;
  /** Checkbox SDT properties (w14:checkbox). */
  readonly checkbox?: SdtCheckboxProperties;
  /** Equation SDT marker. */
  readonly equation?: boolean;
  /** Citation SDT marker. */
  readonly citation?: boolean;
  /** Bibliography SDT marker. */
  readonly bibliography?: boolean;
  /** Temporary (does not persist if content is edited). */
  readonly temporary?: boolean;
  /** Repeating section properties. */
  readonly repeatingSection?: SdtRepeatingSectionProperties;
  /** Repeating section item properties. */
  readonly repeatingSectionItem?: boolean;
}

/** SDT data binding to Custom XML part. */
export interface SdtDataBinding {
  /** XPath expression. */
  readonly xpath: string;
  /** Storage item ID (GUID of the Custom XML part). */
  readonly storeItemId: string;
  /** Prefix mappings for XPath namespaces. */
  readonly prefixMappings?: string;
}

/** SDT checkbox properties (w14:checkbox extension). */
export interface SdtCheckboxProperties {
  /** Is checked. */
  readonly checked?: boolean;
  /** Checked state character (default: ☒ / u+2612). */
  readonly checkedChar?: string;
  /** Checked state font. */
  readonly checkedFont?: string;
  /** Unchecked state character (default: ☐ / u+2610). */
  readonly uncheckedChar?: string;
  /** Unchecked state font. */
  readonly uncheckedFont?: string;
}

/** SDT repeating section properties. */
export interface SdtRepeatingSectionProperties {
  /** Title. */
  readonly title?: string;
  /** Section item name (item displayed in UI). */
  readonly sectionTitle?: string;
  /** Allow insert/delete of repeating items. */
  readonly allowInsertDelete?: boolean;
}

/** SDT content. */
export interface StructuredDocumentTag {
  readonly type: "sdt";
  /** SDT properties. */
  readonly properties?: SdtProperties;
  /**
   * Content (paragraphs, runs, tables, or nested SDTs).
   *
   * Nested SDTs are valid OOXML — repeating sections, for instance, wrap
   * each repeated item in its own inner `<w:sdt>`. The reader preserves
   * the outer/inner structure verbatim and the writer recursively renders
   * children of type `"sdt"`.
   */
  readonly content: readonly (Paragraph | Run | Table | StructuredDocumentTag)[];
}

// =============================================================================
// Legacy Form Fields
// =============================================================================

/** Legacy form field type. */
export type FormFieldType = "text" | "checkBox" | "dropDown";

/** Legacy text form field. */
export interface TextFormField {
  readonly type: "text";
  /** Field name. */
  readonly name?: string;
  /** Default value. */
  readonly default?: string;
  /** Maximum length (0 = unlimited). */
  readonly maxLength?: number;
  /** Text format (e.g. "UPPERCASE", "LOWERCASE", "First capital"). */
  readonly format?: string;
  /** Help text. */
  readonly helpText?: string;
  /** Status bar text. */
  readonly statusText?: string;
  /** Enable/disable field. */
  readonly enabled?: boolean;
}

/** Legacy checkbox form field. */
export interface CheckBoxFormField {
  readonly type: "checkBox";
  /** Field name. */
  readonly name?: string;
  /** Whether checked. */
  readonly checked?: boolean;
  /** Default checked state. */
  readonly default?: boolean;
  /** Size in half-points (auto if omitted). */
  readonly size?: number;
}

/** Legacy dropdown form field. */
export interface DropDownFormField {
  readonly type: "dropDown";
  /** Field name. */
  readonly name?: string;
  /** List entries. */
  readonly entries?: readonly string[];
  /** Default selected index. */
  readonly default?: number;
  /** Help text. */
  readonly helpText?: string;
  /** Status bar text. */
  readonly statusText?: string;
  /** Enable/disable field. */
  readonly enabled?: boolean;
}

/** Any legacy form field. */
export type FormField = TextFormField | CheckBoxFormField | DropDownFormField;

// =============================================================================
// Watermark
// =============================================================================

/** Text watermark properties. */
export interface TextWatermark {
  readonly type: "text";
  /** Watermark text. */
  readonly text: string;
  /** Font name. */
  readonly font?: string;
  /** Font size (half-points). */
  readonly fontSize?: number;
  /** Color (hex, e.g. "C0C0C0"). */
  readonly color?: HexColor;
  /** Semi-transparent. */
  readonly semiTransparent?: boolean;
  /** Rotation in degrees (default -45 for diagonal). */
  readonly rotation?: number;
}

/** Image watermark properties. */
export interface ImageWatermark {
  readonly type: "image";
  /** Relationship ID for the watermark image. */
  readonly rId: string;
  /** Scale percentage (e.g. 100). */
  readonly scale?: number;
  /** Washout effect. */
  readonly washout?: boolean;
}

/** Watermark (text or image). */
export type Watermark = TextWatermark | ImageWatermark;

// =============================================================================
// Document Body Content
// =============================================================================

/** Block-level content that can appear in the document body. */
export type BodyContent =
  | Paragraph
  | Table
  | FloatingImage
  | TableOfContents
  | MathBlock
  | TextBox
  | CheckBox
  | DrawingShape
  | OpaqueDrawing
  | ChartContent
  | ChartExContent
  | AltChunk
  | StructuredDocumentTag;

/** Alternate content chunk (w:altChunk) — embedded HTML/RTF/TXT/etc. in DOCX. */
export interface AltChunk {
  readonly type: "altChunk";
  /** Relationship ID for the alt chunk content. */
  readonly rId: string;
  /** MIME content type of the chunk (e.g. "text/html"). */
  readonly contentType?: string;
  /** Chunk binary data (filled by reader from the archive). */
  readonly data?: Uint8Array;
  /** File name within the archive (e.g. "afchunk.html"). */
  readonly fileName?: string;
}

// =============================================================================
// Charts
// =============================================================================

/** Chart type. */
export type ChartType =
  | "bar"
  | "barStacked"
  | "barPercentStacked"
  | "column"
  | "columnStacked"
  | "columnPercentStacked"
  | "line"
  | "lineStacked"
  | "lineMarked"
  | "pie"
  | "pie3D"
  | "doughnut"
  | "area"
  | "areaStacked"
  | "scatter"
  | "scatterSmooth"
  | "radar"
  | "radarFilled"
  | "bubble"
  | "stock"
  | "surface"
  | "surface3D"
  | "surfaceWireframe"
  | "surfaceWireframe3D";

/** Trendline type for chart series. */
export type ChartTrendlineType =
  | "linear"
  | "exponential"
  | "logarithmic"
  | "polynomial"
  | "power"
  | "movingAvg";

/** Trendline options for a chart series. */
export interface ChartTrendline {
  readonly type: ChartTrendlineType;
  readonly order?: number;
  readonly period?: number;
  readonly displayEquation?: boolean;
  readonly displayRSquared?: boolean;
}

/** Error bar direction. */
export type ChartErrorBarDirection = "x" | "y" | "both";

/** Error bar type. */
export type ChartErrorBarType = "fixedVal" | "percentage" | "stdDev" | "stdErr" | "custom";

/** Error bars options for a chart series. */
export interface ChartErrorBars {
  readonly direction: ChartErrorBarDirection;
  readonly type: ChartErrorBarType;
  readonly value?: number;
}

/** A data series in a chart. */
export interface ChartSeries {
  readonly name: string;
  readonly categories: readonly string[];
  readonly values: readonly number[];
  readonly color?: HexColor;
  readonly pointColors?: readonly HexColor[];
  readonly showDataLabels?: boolean;
  readonly trendline?: ChartTrendline;
  readonly errorBars?: ChartErrorBars;
  /** Per-series chart type override (for combo charts). */
  readonly chartType?: ChartType;
  /** Plot on secondary axis (for combo charts). */
  readonly plotOnSecondaryAxis?: boolean;
}

/** Chart legend position. */
export type ChartLegendPosition = "b" | "l" | "r" | "t" | "tr" | "none";

/** Data label position. */
export type ChartDataLabelPosition = "outsideEnd" | "center" | "insideEnd" | "bestFit";

/** Data labels options for a chart. */
export interface ChartDataLabels {
  readonly showValue?: boolean;
  readonly showCategory?: boolean;
  readonly showSerName?: boolean;
  readonly showPercent?: boolean;
  readonly position?: ChartDataLabelPosition;
}

/** Chart axis configuration. */
export interface ChartAxis {
  readonly title?: string;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly numberFormat?: string;
  readonly hidden?: boolean;
}

/** Chart definition. */
export interface Chart {
  readonly type: ChartType;
  readonly title?: string;
  readonly series: readonly ChartSeries[];
  readonly legend?: ChartLegendPosition;
  readonly dataLabels?: ChartDataLabels;
  readonly categoryAxis?: ChartAxis;
  readonly valueAxis?: ChartAxis;
  readonly secondaryValueAxis?: ChartAxis;
  readonly plotAreaColor?: HexColor;
  readonly chartAreaColor?: HexColor;
  readonly view3d?: boolean;
  readonly style?: number;
  readonly width?: Emu;
  readonly height?: Emu;
  /** Secondary chart type for combo charts (legacy approach). */
  readonly secondaryType?: ChartType;
  /** Secondary series for combo chart (plotted on secondary axis). */
  readonly secondarySeries?: readonly ChartSeries[];
  /** Whether to embed chart data as an xlsx workbook (for full editing in Word). */
  readonly embedWorkbook?: boolean;
}

/** Chart content placed in a document. */
export interface ChartContent {
  readonly type: "chart";
  readonly chart: Chart;
  readonly altText?: string;
  readonly name?: string;
}

/**
 * ChartEx content block (cx: namespace, Office 2016+ chart types).
 * Supports: sunburst, treemap, waterfall, funnel, histogram, pareto, boxWhisker, regionMap.
 */
export interface ChartExContent {
  readonly type: "chartEx";
  /** Pre-rendered ChartEx XML (cx:chartSpace). Generated via excel-bridge. */
  readonly chartExXml: string;
  /** Parsed structured data extracted from the ChartEx XML. */
  readonly data?: ChartExData;
  /** Alt text for accessibility. */
  readonly altText?: string;
  /** Drawing name. */
  readonly name?: string;
  /** Width in EMU. Default: 5486400 (6 inches). */
  readonly width?: Emu;
  /** Height in EMU. Default: 3657600 (4 inches). */
  readonly height?: Emu;
}

/** Structured data extracted from a ChartEx (cx: namespace) XML. */
export interface ChartExData {
  /** Chart type (sunburst, treemap, waterfall, funnel, histogram, pareto, boxWhisker, regionMap). */
  readonly chartType: string;
  /** Chart title. */
  readonly title?: string;
  /** Series data. */
  readonly series: readonly ChartExSeriesData[];
}

/** A single data series within a ChartEx chart. */
export interface ChartExSeriesData {
  /** Series name/label. */
  readonly name?: string;
  /** String categories. */
  readonly categories?: readonly string[];
  /** Numeric values. */
  readonly values?: readonly number[];
}

// =============================================================================
// Styles
// =============================================================================

/** Style type. */
export type StyleType = "paragraph" | "character" | "table" | "numbering";

/** Table style conditional format type. */
export type TableStyleConditionType =
  | "firstRow"
  | "lastRow"
  | "firstColumn"
  | "lastColumn"
  | "oddRowBanding"
  | "evenRowBanding"
  | "oddColumnBanding"
  | "evenColumnBanding"
  | "topLeftCell"
  | "topRightCell"
  | "bottomLeftCell"
  | "bottomRightCell";

/** Table style conditional formatting. */
export interface TableStyleConditionalFormat {
  /** Condition type. */
  readonly type: TableStyleConditionType;
  /** Paragraph properties for this condition. */
  readonly paragraphProperties?: ParagraphProperties;
  /** Run properties for this condition. */
  readonly runProperties?: RunProperties;
  /** Table properties for this condition. */
  readonly tableProperties?: TableProperties;
  /** Table row properties for this condition. */
  readonly rowProperties?: TableRowProperties;
  /** Table cell properties for this condition. */
  readonly cellProperties?: TableCellProperties;
}

/** Style definition. */
export interface StyleDef {
  /** Style type. */
  readonly type: StyleType;
  /** Unique style ID. */
  readonly styleId: string;
  /** Display name. */
  readonly name: string;
  /** Parent style ID. */
  readonly basedOn?: string;
  /** Next paragraph style ID. */
  readonly next?: string;
  /** Linked style ID. */
  readonly link?: string;
  /** Show in Quick Styles gallery. */
  readonly qFormat?: boolean;
  /** UI priority (lower = higher priority). */
  readonly uiPriority?: number;
  /** Is this the default style for its type. */
  readonly isDefault?: boolean;
  /** Paragraph properties. */
  readonly paragraphProperties?: ParagraphProperties;
  /** Run properties. */
  readonly runProperties?: RunProperties;
  /** Table properties (for table styles). */
  readonly tableProperties?: TableProperties;
  /** Outline level (0-8 for Heading styles). */
  readonly outlineLevel?: number;
  /** Semi-hidden (not shown in style gallery). */
  readonly semiHidden?: boolean;
  /** Unhide when used. */
  readonly unhideWhenUsed?: boolean;
  /** Custom (user-created) style (w:customStyle attribute). */
  readonly customStyle?: boolean;
  /** Hidden (not shown anywhere). */
  readonly hidden?: boolean;
  /** Lock style (prevent modifications). */
  readonly locked?: boolean;
  /** Enable auto-redefining the style when used. */
  readonly autoRedefine?: boolean;
  /** Table style conditional formats. */
  readonly tableStyleConditions?: readonly TableStyleConditionalFormat[];
}

/** Document defaults for styles.xml. */
export interface DocDefaults {
  /** Default run properties. */
  readonly runProperties?: RunProperties;
  /** Default paragraph properties. */
  readonly paragraphProperties?: ParagraphProperties;
}

// =============================================================================
// Numbering / Lists
// =============================================================================

/** Number format type (ST_NumberFormat). */
export type NumberFormat =
  | "decimal"
  | "upperRoman"
  | "lowerRoman"
  | "upperLetter"
  | "lowerLetter"
  | "ordinal"
  | "cardinalText"
  | "ordinalText"
  | "hex"
  | "chicago"
  | "ideographDigital"
  | "japaneseCounting"
  | "aiueo"
  | "iroha"
  | "decimalFullWidth"
  | "decimalHalfWidth"
  | "japaneseLegal"
  | "japaneseDigitalTenThousand"
  | "decimalEnclosedCircle"
  | "decimalFullWidth2"
  | "aiueoFullWidth"
  | "irohaFullWidth"
  | "decimalZero"
  | "bullet"
  | "ganada"
  | "chosung"
  | "decimalEnclosedFullstop"
  | "decimalEnclosedParen"
  | "decimalEnclosedCircleChinese"
  | "ideographEnclosedCircle"
  | "ideographTraditional"
  | "ideographZodiac"
  | "ideographZodiacTraditional"
  | "taiwaneseCounting"
  | "ideographLegalTraditional"
  | "taiwaneseCountingThousand"
  | "taiwaneseDigital"
  | "chineseCounting"
  | "chineseLegalSimplified"
  | "chineseCountingThousand"
  | "koreanDigital"
  | "koreanCounting"
  | "koreanLegal"
  | "koreanDigital2"
  | "vietnameseCounting"
  | "russianLower"
  | "russianUpper"
  | "none"
  | "numberInDash"
  | "hebrew1"
  | "hebrew2"
  | "arabicAlpha"
  | "arabicAbjad"
  | "hindiVowels"
  | "hindiConsonants"
  | "hindiNumbers"
  | "hindiCounting"
  | "thaiLetters"
  | "thaiNumbers"
  | "thaiCounting"
  | "bahtText"
  | "dollarText"
  | "custom";

/** Level justification. */
export type LevelJustification = "left" | "center" | "right";

/** Level suffix (character after number). */
export type LevelSuffix = "tab" | "space" | "nothing";

/** Numbering level definition. */
export interface NumberingLevel {
  /** Level index (0-8). */
  readonly level: number;
  /** Start value. */
  readonly start?: number;
  /** Number format. */
  readonly format: NumberFormat;
  /** Level text (e.g. "%1." or "%1.%2"). For bullet, a symbol character. */
  readonly text: string;
  /** Paragraph style ID linked to this level. */
  readonly paragraphStyle?: string;
  /** Justification. */
  readonly justification?: LevelJustification;
  /** Paragraph properties for this level. */
  readonly paragraphProperties?: ParagraphProperties;
  /** Run properties for the numbering symbol. */
  readonly runProperties?: RunProperties;
  /** Character after number: tab, space, or nothing. */
  readonly suffix?: LevelSuffix;
  /** Legal numbering style (override all lower levels to decimal). */
  readonly isLegalNumberingStyle?: boolean;
  /** Restart after which level (-1 = never, 0 = level 0 restarts this). */
  readonly restartAfterLevel?: number;
  /** Picture bullet ID (references w:numPicBullet in numbering.xml). */
  readonly picBulletId?: number;
}

/** Picture bullet definition (for numbering with image bullets). */
export interface NumPicBullet {
  /** Unique ID (referenced by NumberingLevel.picBulletId). */
  readonly id: number;
  /** Relationship ID to the image in numbering.xml.rels. */
  readonly rId?: string;
  /** Width in EMU. */
  readonly width?: Emu;
  /** Height in EMU. */
  readonly height?: Emu;
  /** Raw VML fallback XML. */
  readonly rawVmlXml?: string;
}

/** Multi-level type. */
export type MultiLevelType = "singleLevel" | "multilevel" | "hybridMultilevel";

/** Abstract numbering definition (template). */
export interface AbstractNumbering {
  /** Unique abstract numbering ID. */
  readonly abstractNumId: number;
  /** Multi-level type. */
  readonly multiLevelType?: MultiLevelType;
  /** Level definitions. */
  readonly levels: readonly NumberingLevel[];
  /** Numbering style link. */
  readonly numStyleLink?: string;
  /** Style link. */
  readonly styleLink?: string;
}

/** Level override for a numbering instance. */
export interface LevelOverride {
  readonly level: number;
  readonly startOverride?: number;
  /** Level definition override. */
  readonly levelDef?: NumberingLevel;
}

/** Numbering instance (referenced by paragraphs). */
export interface NumberingInstance {
  /** Unique numbering ID (referenced by paragraphs via numId). */
  readonly numId: number;
  /** Abstract numbering ID this instance is based on. */
  readonly abstractNumId: number;
  /** Level overrides. */
  readonly overrides?: readonly LevelOverride[];
}

// =============================================================================
// Headers & Footers
// =============================================================================

/** Header/footer content (same structure as document body). */
export interface HeaderFooterContent {
  /** Content paragraphs. */
  readonly children: readonly (Paragraph | Table)[];
}

/** Header definition. */
export interface HeaderDef {
  /** Header content. */
  readonly content: HeaderFooterContent;
  /** Relationship ID (assigned during packaging). */
  rId?: string;
}

/** Footer definition. */
export interface FooterDef {
  /** Footer content. */
  readonly content: HeaderFooterContent;
  /** Relationship ID (assigned during packaging). */
  rId?: string;
}

// =============================================================================
// Footnotes & Endnotes
// =============================================================================

/** Footnote/endnote numbering format. */
export type NoteNumberFormat =
  | "decimal"
  | "upperRoman"
  | "lowerRoman"
  | "upperLetter"
  | "lowerLetter"
  | "chicago"
  | "bullet";

/** Footnote/endnote numbering restart. */
export type NoteNumberRestart = "continuous" | "eachSect" | "eachPage";

/** Footnote properties (w:footnotePr). */
export interface FootnoteProperties {
  /** Numbering format. */
  readonly numFmt?: NoteNumberFormat;
  /** Starting number. */
  readonly numStart?: number;
  /** Numbering restart rule. */
  readonly numRestart?: NoteNumberRestart;
  /** Footnote position: pageBottom or beneathText. */
  readonly position?: "pageBottom" | "beneathText";
}

/** Endnote properties (w:endnotePr). */
export interface EndnoteProperties {
  /** Numbering format. */
  readonly numFmt?: NoteNumberFormat;
  /** Starting number. */
  readonly numStart?: number;
  /** Numbering restart rule. */
  readonly numRestart?: NoteNumberRestart;
  /** Endnote position: sectEnd or docEnd. */
  readonly position?: "sectEnd" | "docEnd";
}

/** Footnote definition. */
/** Footnote/endnote note type. */
export type NoteType = "normal" | "separator" | "continuationSeparator" | "continuationNotice";

export interface FootnoteDef {
  /** Unique footnote ID (starts from 1; -1 and 0 are reserved for separators). */
  readonly id: number;
  /** Note type (default: "normal"). */
  readonly type?: NoteType;
  /** Footnote content. */
  readonly content: readonly Paragraph[];
}

/** Endnote definition. */
export interface EndnoteDef {
  /** Unique endnote ID (starts from 1; -1 and 0 are reserved for separators). */
  readonly id: number;
  /** Note type (default: "normal"). */
  readonly type?: NoteType;
  /** Endnote content. */
  readonly content: readonly Paragraph[];
}

// =============================================================================
// Images / Media
// =============================================================================

/** Image media type. */
export type ImageMediaType =
  | "png"
  | "jpeg"
  | "gif"
  | "bmp"
  | "tiff"
  | "svg"
  | "webp"
  | "emf"
  | "wmf";

/** Image definition in the media collection. */
export interface ImageDef {
  /** Image data. */
  readonly data: Uint8Array;
  /** MIME media type. */
  readonly mediaType: ImageMediaType;
  /** File name within word/media/ (e.g. "image1.png"). */
  readonly fileName: string;
  /** Relationship ID (assigned during packaging). */
  rId?: string;
  /**
   * Additional relationship IDs under which this same media file is referenced
   * by other parts (typically header/footer .rels using their own local id
   * space). Populated by the reader so that the packager can rebuild each
   * part's local .rels even when callers and parts use different rIds for the
   * same physical media file.
   */
  readonly aliasRIds?: readonly string[];
  /** SVG fallback image data (PNG). Required for SVG images in Word. */
  readonly fallbackData?: Uint8Array;
}

// =============================================================================
// Font Table
// =============================================================================

/** Font family classification. */
export type FontFamily = "roman" | "swiss" | "modern" | "decorative" | "script" | "auto";

/** Font pitch. */
export type FontPitch = "fixed" | "variable" | "default";

/** Font definition for fontTable.xml. */
export interface FontDef {
  readonly name: string;
  readonly panose1?: string;
  readonly charset?: string;
  readonly family?: FontFamily;
  readonly pitch?: FontPitch;
  /** Signature bytes (usb0-3, csb0-1). */
  readonly sig?: {
    readonly usb0?: string;
    readonly usb1?: string;
    readonly usb2?: string;
    readonly usb3?: string;
    readonly csb0?: string;
    readonly csb1?: string;
  };
  /** Embedded regular font relationship ID. */
  readonly embedRegular?: string;
  /** Embedded regular font key (GUID used for obfuscation). */
  readonly embedRegularKey?: string;
  /** Embedded bold font relationship ID. */
  readonly embedBold?: string;
  /** Embedded bold font key. */
  readonly embedBoldKey?: string;
  /** Embedded italic font relationship ID. */
  readonly embedItalic?: string;
  /** Embedded italic font key. */
  readonly embedItalicKey?: string;
  /** Embedded bold-italic font relationship ID. */
  readonly embedBoldItalic?: string;
  /** Embedded bold-italic font key. */
  readonly embedBoldItalicKey?: string;
}

/** Embedded font binary data. */
export interface EmbeddedFont {
  /** Relationship ID (as referenced in FontDef.embedRegular etc.). */
  readonly rId: string;
  /** Font data — can be raw TTF/OTF or obfuscated ODTTF. */
  readonly data: Uint8Array;
  /** GUID key used for obfuscation (if obfuscated). When set, data is ODTTF (first 32 bytes XORed with GUID). */
  readonly fontKey?: string;
  /** File name inside word/fonts/ (e.g. "font1.odttf"). */
  readonly fileName: string;
}

// =============================================================================
// Document Settings
// =============================================================================

/** Document protection type. */
export type ProtectionType = "none" | "readOnly" | "comments" | "trackedChanges" | "forms";

/** Hyphenation settings. */
export interface HyphenationSettings {
  /** Automatic hyphenation. */
  readonly autoHyphenation?: boolean;
  /** Hyphenation zone (distance from right margin in twips). */
  readonly hyphenationZone?: Twips;
  /** Maximum consecutive hyphenated lines (0 = unlimited). */
  readonly consecutiveHyphenLimit?: number;
  /** Do not hyphenate words in all caps. */
  readonly doNotHyphenateCaps?: boolean;
}

/** Document settings. */
export interface DocumentSettings {
  /** Zoom percentage. */
  readonly zoom?: number;
  /** Default tab stop distance in twips. */
  readonly defaultTabStop?: Twips;
  /** Character spacing compression control. */
  readonly characterSpacingControl?:
    | "doNotCompress"
    | "compressPunctuation"
    | "compressPunctuationAndJapaneseKana";
  /** Compatibility mode version. */
  readonly compatibilityMode?: number;
  /** Even and odd page headers/footers. */
  readonly evenAndOddHeaders?: boolean;
  /** Track changes / revisions enabled. */
  readonly trackRevisions?: boolean;
  /** Document protection. */
  readonly documentProtection?: {
    readonly type?: ProtectionType;
    readonly edit?: string;
    readonly enforcement?: boolean;
    readonly formatting?: boolean;
    /** Hash algorithm (e.g. "SHA-256"). */
    readonly hashAlgorithm?: string;
    /** Base64-encoded hash value. */
    readonly hashValue?: string;
    /** Base64-encoded salt value. */
    readonly saltValue?: string;
    /** Number of hash iterations. */
    readonly spinCount?: number;
  };
  /** Hyphenation settings. */
  readonly hyphenation?: HyphenationSettings;
  /** Auto-hyphenation. */
  readonly autoHyphenation?: boolean;
  /** Mirror margins. */
  readonly mirrorMargins?: boolean;
  /** Gutter at top. */
  readonly gutterAtTop?: boolean;
  /** Display background colors/images in print layout. */
  readonly displayBackgroundShape?: boolean;
  /** Update fields on open. */
  readonly updateFieldsOnOpen?: boolean;
  /** Document variables (key-value pairs). */
  readonly docVars?: ReadonlyMap<string, string>;
  /** Document-level footnote properties. */
  readonly footnoteProperties?: FootnoteProperties;
  /** Document-level endnote properties. */
  readonly endnoteProperties?: EndnoteProperties;
  /** All w:compatSetting entries (except compatibilityMode which has its own field). */
  readonly compatSettings?: ReadonlyArray<CompatSetting>;
  /** Legacy w:compat child elements (e.g. w:useFELayout, w:balanceSingleByteDoubleByteWidth). */
  readonly compatFlags?: ReadonlyArray<CompatFlag>;
  /** Mail merge settings (opaque). */
  readonly mailMergeRawXml?: string;
  /** Write protection (separate from documentProtection). */
  readonly writeProtection?: {
    readonly recommended?: boolean;
    readonly algorithmName?: string;
    readonly hashValue?: string;
    readonly saltValue?: string;
    readonly spinCount?: number;
  };
  /** RSID revision save IDs (raw). */
  readonly rsids?: {
    readonly rsidRoot?: string;
    readonly rsid?: readonly string[];
  };
  /** Decimal symbol for numbers (e.g. "."). */
  readonly decimalSymbol?: string;
  /** List separator for fields (e.g. ","). */
  readonly listSeparator?: string;
  /** Do not track move revisions separately. */
  readonly doNotTrackMoves?: boolean;
  /** Do not track formatting changes. */
  readonly doNotTrackFormatting?: boolean;
  /** Save subset fonts (only used glyphs). */
  readonly saveSubsetFonts?: boolean;
  /** Borders should not surround header area. */
  readonly bordersDoNotSurroundHeader?: boolean;
  /** Borders should not surround footer area. */
  readonly bordersDoNotSurroundFooter?: boolean;
  /** Theme language defaults (for themeFontLang). */
  readonly themeFontLang?: {
    readonly val?: string;
    readonly eastAsia?: string;
    readonly bidi?: string;
  };
  /** Click-and-type default style ID. */
  readonly clickAndTypeStyle?: string;
  /** Disable punctuation kerning. */
  readonly noPunctuationKerning?: boolean;
  /** Style pane format filter (bitmask hex string). */
  readonly stylePaneFormatFilter?: string;
  /** Style pane sort method. */
  readonly stylePaneSortMethod?: string;
  /** Do not demote first line of Asian text. */
  readonly doNotDemoteAsianTextFirstLine?: boolean;
}

/** A w:compatSetting entry in settings.xml. */
export interface CompatSetting {
  readonly name: string;
  readonly uri: string;
  readonly val: string;
}

/** A legacy w:compat child flag. */
export interface CompatFlag {
  /** Element name without namespace (e.g. "useFELayout"). */
  readonly name: string;
  /** Optional w:val attribute. */
  readonly val?: string;
}

// =============================================================================
// Document Background
// =============================================================================

/** Document background (w:background). */
export interface DocumentBackground {
  /** Background color (hex). */
  readonly color?: HexColor;
  /** Theme color. */
  readonly themeColor?: string;
  /** Theme shade. */
  readonly themeShade?: string;
  /** Theme tint. */
  readonly themeTint?: string;
}

// =============================================================================
// Custom Properties
// =============================================================================

/** Custom property value type. */
export type CustomPropertyValue =
  | { readonly type: "string"; readonly value: string }
  | { readonly type: "number"; readonly value: number }
  | { readonly type: "boolean"; readonly value: boolean }
  | { readonly type: "date"; readonly value: Date };

/** Custom document property. */
export interface CustomProperty {
  /** Property name. */
  readonly name: string;
  /** Property value. */
  readonly value: CustomPropertyValue;
}

// =============================================================================
// Core Properties (docProps/core.xml)
// =============================================================================

/** Dublin Core metadata. */
export interface CoreProperties {
  readonly title?: string;
  readonly subject?: string;
  readonly creator?: string;
  readonly description?: string;
  readonly keywords?: string;
  readonly lastModifiedBy?: string;
  readonly revision?: string;
  readonly created?: Date;
  readonly modified?: Date;
  readonly category?: string;
}

// =============================================================================
// App Properties (docProps/app.xml)
// =============================================================================

/** Application properties. */
export interface AppProperties {
  readonly application?: string;
  readonly appVersion?: string;
  readonly pages?: number;
  readonly words?: number;
  readonly characters?: number;
  readonly lines?: number;
  readonly paragraphs?: number;
  readonly company?: string;
  readonly manager?: string;
}

// =============================================================================
// Theme
// =============================================================================

/** Standard OOXML theme color names. */
export type ThemeColorName =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

/** Theme color scheme — maps color names to hex RGB values. */
export interface ThemeColorScheme {
  readonly name: string;
  readonly colors: Readonly<Record<ThemeColorName, HexColor>>;
}

/** Theme font definition (major or minor). */
export interface ThemeFont {
  /** Latin typeface. */
  readonly latin: string;
  /** East Asian typeface (CJK). */
  readonly eastAsia?: string;
  /** Complex Script typeface (Arabic, Hebrew, etc.). */
  readonly complexScript?: string;
  /** Supplemental fonts: script name → typeface. */
  readonly supplementalFonts?: Readonly<Record<string, string>>;
}

/** Theme font scheme. */
export interface ThemeFontScheme {
  readonly name: string;
  /** Major (heading) font — legacy simple form. */
  readonly majorFont: string;
  /** Minor (body) font — legacy simple form. */
  readonly minorFont: string;
  /** Full major font with EA/CS/supplemental. */
  readonly major?: ThemeFont;
  /** Full minor font with EA/CS/supplemental. */
  readonly minor?: ThemeFont;
}

/** Theme format scheme — preserves the opaque XML for fillStyleLst/lnStyleLst/effectStyleLst/bgFillStyleLst. */
export interface ThemeFormatScheme {
  readonly name: string;
  /** Raw XML for a:fmtScheme children (for round-trip preservation). */
  readonly rawXml?: string;
}

/** Parsed document theme. */
export interface DocumentTheme {
  readonly name?: string;
  readonly colorScheme: ThemeColorScheme;
  readonly fontScheme: ThemeFontScheme;
  /** Format scheme (effect/fill/line/bg styles). */
  readonly formatScheme?: ThemeFormatScheme;
  /** Raw XML of extra <a:extLst> element inside <a:theme>, for round-trip. */
  readonly extLstXml?: string;
}

// =============================================================================
// Document Model (Top-Level)
// =============================================================================

/** Complete DOCX document model. */
/** Web settings (word/webSettings.xml). */
export interface WebSettings {
  /** Optimize for specific browser. */
  readonly optimizeForBrowser?: {
    readonly target?: string;
    readonly majorVersion?: number;
  };
  /** Allow PNG images. */
  readonly allowPng?: boolean;
  /** Rely on VML (for older browsers). */
  readonly relyOnVml?: boolean;
  /** Do not save file as single HTML. */
  readonly doNotSaveAsSingleFile?: boolean;
  /** Do not organize into folders. */
  readonly doNotOrganizeInFolder?: boolean;
  /** Use target machine type (for embedded resources). */
  readonly useTargetMachineType?: boolean;
  /** Raw XML (for unsupported settings, round-trip preservation). */
  readonly rawXml?: string;
}

/** Person info (word/people.xml) — co-author/commenter metadata. */
export interface PersonInfo {
  /** Author name (w15:person w15:author). */
  readonly author: string;
  /** Presence info (optional). */
  readonly presenceInfo?: {
    readonly providerId?: string;
    readonly userId?: string;
  };
}

/**
 * Document type discriminator.
 * - "document" — standard .docx
 * - "template" — .dotx (Word template)
 * - "macroEnabledDocument" — .docm (macro-enabled)
 * - "macroEnabledTemplate" — .dotm (macro-enabled template)
 */
export type DocxDocumentType =
  | "document"
  | "template"
  | "macroEnabledDocument"
  | "macroEnabledTemplate";

export interface DocxDocument {
  /**
   * Document type. Determines the content type used in the package.
   * Default: "document" (standard .docx).
   */
  readonly docType?: DocxDocumentType;

  /** Document body content. */
  readonly body: readonly BodyContent[];

  /** Final section properties (for the last section). */
  readonly sectionProperties?: SectionProperties;

  /** Style definitions. */
  readonly styles?: readonly StyleDef[];
  /** Document defaults. */
  readonly docDefaults?: DocDefaults;

  /** Abstract numbering definitions. */
  readonly abstractNumberings?: readonly AbstractNumbering[];
  /** Numbering instances. */
  readonly numberingInstances?: readonly NumberingInstance[];
  /** Picture bullet definitions (for numbering with images). */
  readonly numPicBullets?: readonly NumPicBullet[];

  /** Headers (keyed by internal name, e.g. "default", "first", "even"). */
  readonly headers?: ReadonlyMap<string, HeaderDef>;
  /** Footers. */
  readonly footers?: ReadonlyMap<string, FooterDef>;

  /** Footnotes. */
  readonly footnotes?: readonly FootnoteDef[];
  /** Endnotes. */
  readonly endnotes?: readonly EndnoteDef[];

  /** Images / media. */
  readonly images?: readonly ImageDef[];

  /** Font table entries. */
  readonly fonts?: readonly FontDef[];

  /** Embedded font binaries (stored in word/fonts/). */
  readonly embeddedFonts?: readonly EmbeddedFont[];

  /** Custom XML parts (for data binding with SDT). */
  readonly customXmlParts?: readonly CustomXmlPart[];

  /** Document settings. */
  readonly settings?: DocumentSettings;

  /** Web settings (word/webSettings.xml). */
  readonly webSettings?: WebSettings;

  /** Document thumbnail for file manager previews. */
  readonly thumbnail?: {
    readonly contentType: "image/jpeg" | "image/x-wmf" | "image/png";
    readonly data: Uint8Array;
  };

  /** Authors/commenters metadata (word/people.xml). */
  readonly people?: readonly PersonInfo[];

  /** Core properties (metadata). */
  readonly coreProperties?: CoreProperties;
  /** Application properties. */
  readonly appProperties?: AppProperties;

  /** Comments. */
  readonly comments?: readonly CommentDef[];

  /** Document background. */
  readonly background?: DocumentBackground;

  /** Custom properties. */
  readonly customProperties?: readonly CustomProperty[];

  /** Watermark (rendered in default header). */
  readonly watermark?: Watermark;

  /** Document theme (from theme1.xml). */
  readonly theme?: DocumentTheme;

  /** Opaque (unrecognized) ZIP parts preserved for round-trip fidelity. */
  readonly opaqueParts?: readonly OpaquePart[];

  /** VBA project binary (word/vbaProject.bin) for .docm/.dotm round-trip. */
  readonly vbaProject?: Uint8Array;
}

// =============================================================================
// Opaque Parts (Round-trip preservation)
// =============================================================================

/** A relationship entry preserved for round-trip. */
export interface OpaqueRelationship {
  /** Relationship ID (e.g. "rId5"). */
  readonly id: string;
  /** Relationship type URI. */
  readonly type: string;
  /** Target path (relative to source part). */
  readonly target: string;
  /**
   * Target mode for the relationship. Standard OPC values are `"External"`
   * and `"Internal"` (with `"Internal"` typically omitted on the wire). The
   * field is typed as `string` rather than a union so that round-trip
   * preserves any non-standard value present in the source document — the
   * writer surfaces this verbatim into the output `.rels` file.
   */
  readonly targetMode?: string;
}

/** Custom XML part (for data binding). Stored in word/customXml/. */
export interface CustomXmlPart {
  /** Unique item ID (GUID). Referenced by SdtDataBinding.storeItemId. */
  readonly itemId: string;
  /** Raw XML content (word/customXml/item*.xml). */
  readonly xmlContent: string;
  /** Schema references (XML namespaces used). */
  readonly schemaReferences?: readonly string[];
  /** File name (e.g. "item1.xml"). */
  readonly fileName: string;
}

/** An opaque ZIP part preserved for round-trip fidelity. */
export interface OpaquePart {
  /** Part path in the ZIP (e.g. "word/charts/chart1.xml"). */
  readonly path: string;
  /** Raw content of the part. */
  readonly data: Uint8Array;
  /** Content type (from [Content_Types].xml). */
  readonly contentType?: string;
  /** Relationships of this part (from its .rels file). */
  readonly relationships?: readonly OpaqueRelationship[];
}

/** An opaque drawing element (e.g. chart) preserved in a paragraph. */
export interface OpaqueDrawing {
  readonly type: "opaqueDrawing";
  /** Raw XML of the w:drawing element. */
  readonly rawXml: string;
  /** Relationship IDs referenced within this drawing (for part resolution). */
  readonly referencedRIds: readonly string[];
}

// =============================================================================
// Builder Options
// =============================================================================

/** Options for creating a DOCX file. */
export interface DocxOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
}

// =============================================================================
// Asset References (currently unused; reserved for future package-model API)
// =============================================================================

/** A reference to a media asset (image, chart, etc.) by logical ID. */
export interface AssetRef {
  /** Logical asset ID (assigned by builder, resolved to rId at package time). */
  readonly assetId: string;
  /** Asset type hint. */
  readonly type: "image" | "chart" | "oleObject" | "font" | "xlsx";
}

/** A media asset registered in the document. */
export interface MediaAsset {
  /** Logical ID for referencing. */
  readonly id: string;
  /** MIME content type. */
  readonly contentType: string;
  /** Raw binary data. */
  readonly data: Uint8Array;
  /** Suggested file name (without path). */
  readonly fileName: string;
}
