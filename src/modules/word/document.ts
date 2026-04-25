/**
 * DOCX Module - Document Builder
 *
 * High-level fluent API for constructing DOCX documents programmatically.
 * Provides convenience methods for common operations including comments,
 * track changes, TOC, math, text boxes, checkboxes, and custom properties.
 */

import { packageDocx } from "./docx-packager";
import { readDocx } from "./docx-reader";
import { bytesToBase64 } from "./internal-utils";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphProperties,
  ParagraphChild,
  Run,
  RunProperties,
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TableRowProperties,
  TableCellProperties,
  SectionProperties,
  StyleDef,
  DocDefaults,
  AbstractNumbering,
  NumberingInstance,
  BookmarkStart,
  Hyperlink,
  HeaderDef,
  FooterDef,
  HeaderFooterContent,
  FootnoteDef,
  EndnoteDef,
  ImageDef,
  ImageMediaType,
  CoreProperties,
  AppProperties,
  DocumentSettings,
  FontDef,
  Border,
  TableWidth,
  TableBorders,
  Emu,
  Twips,
  CommentDef,
  DocumentBackground,
  CustomProperty,
  CustomPropertyValue,
  TableOfContents,
  MathBlock,
  MathContent,
  CheckBox,
  RevisionInfo,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  MoveRangeMarker,
  Chart,
  ChartSeries,
  ChartAxis,
  ChartLegendPosition,
  ChartContent,
  CommentRangeStart,
  CommentRangeEnd,
  CommentReference,
  FloatingImage,
  InlineImageContent,
  StructuredDocumentTag,
  SdtProperties,
  Watermark,
  DocumentTheme,
  ColorSpec,
  HexColor,
  DrawingShape,
  RunContent,
  PositionalTabAlignment,
  PositionalTabRelativeTo,
  PositionalTabLeader,
  RubyProperties
} from "./types";

// =============================================================================
// Helper Builders
// =============================================================================

/** Create a text run. */
export function text(content: string, properties?: RunProperties): Run {
  return { properties, content: [{ type: "text", text: content }] };
}

/** Create a bold text run. */
export function bold(content: string, properties?: Omit<RunProperties, "bold">): Run {
  return text(content, { ...properties, bold: true });
}

/** Create an italic text run. */
export function italic(content: string, properties?: Omit<RunProperties, "italic">): Run {
  return text(content, { ...properties, italic: true });
}

/** Create a run with a page break. */
export function pageBreak(): Run {
  return { content: [{ type: "break", breakType: "page" }] };
}

/** Create a run with a line break. */
export function lineBreak(): Run {
  return { content: [{ type: "break" }] };
}

/** Create a run with a column break. */
export function columnBreak(): Run {
  return { content: [{ type: "break", breakType: "column" }] };
}

/** Create a tab run. */
export function tab(): Run {
  return { content: [{ type: "tab" }] };
}

/** Create a positional tab (w:ptab). */
export function positionalTab(options: {
  alignment: PositionalTabAlignment;
  relativeTo: PositionalTabRelativeTo;
  leader?: PositionalTabLeader;
}): Run {
  return {
    content: [
      {
        type: "ptab",
        alignment: options.alignment,
        relativeTo: options.relativeTo,
        leader: options.leader
      }
    ]
  };
}

/**
 * Create a ruby (phonetic guide) run — e.g. Japanese furigana or Chinese pinyin.
 *
 * @param baseText - The main text being annotated.
 * @param rubyText - The phonetic text shown above the base.
 * @param properties - Optional ruby properties (alignment, font size, language).
 */
export function ruby(
  baseText: string | Run | readonly Run[],
  rubyText: string | Run | readonly Run[],
  properties?: RubyProperties
): Run {
  const normalize = (v: string | Run | readonly Run[]): readonly Run[] => {
    if (typeof v === "string") {
      return [text(v)];
    }
    if (Array.isArray(v)) {
      return v as readonly Run[];
    }
    return [v as Run];
  };
  return {
    content: [
      {
        type: "ruby",
        properties,
        baseText: normalize(baseText),
        rubyText: normalize(rubyText)
      }
    ]
  };
}

/** Create a carriage return run. */
export function carriageReturn(): Run {
  return { content: [{ type: "carriageReturn" }] };
}

/** Create a no-break hyphen run. */
export function noBreakHyphen(): Run {
  return { content: [{ type: "noBreakHyphen" }] };
}

/** Create a soft hyphen run. */
export function softHyphen(): Run {
  return { content: [{ type: "softHyphen" }] };
}

/** Create a run with a field code. */
export function field(instruction: string, cachedValue?: string): Run {
  return { content: [{ type: "field", instruction, cachedValue }] };
}

/** Create a PAGE field (current page number). */
export function pageNumberField(cachedValue?: string): Run {
  return field(" PAGE ", cachedValue ?? "1");
}

/** Create a NUMPAGES field (total page count). */
export function totalPagesField(cachedValue?: string): Run {
  return field(" NUMPAGES ", cachedValue ?? "1");
}

/** Create a SECTIONPAGES field (pages in section). */
export function sectionPagesField(cachedValue?: string): Run {
  return field(" SECTIONPAGES ", cachedValue ?? "1");
}

/** Create a SECTION field (current section number). */
export function sectionField(cachedValue?: string): Run {
  return field(" SECTION ", cachedValue ?? "1");
}

/** Create a DATE field. */
export function dateField(format?: string, cachedValue?: string): Run {
  const fmt = format ?? "yyyy-MM-dd";
  return field(` DATE \\@ "${fmt}" `, cachedValue);
}

/** Create a SEQ (sequence) field for numbering figures, tables, etc. */
export function sequenceField(
  identifier: string,
  options?: { cachedValue?: string; hide?: boolean }
): Run {
  let instruction = ` SEQ ${identifier} `;
  if (options?.hide) {
    instruction += "\\h ";
  }
  return field(instruction, options?.cachedValue);
}

/** Create a TIME field (current time). */
export function timeField(format?: string, cachedValue?: string): Run {
  const fmt = format ?? "HH:mm:ss";
  return field(` TIME \\@ "${fmt}" `, cachedValue);
}

/** Create an AUTHOR field. */
export function authorField(cachedValue?: string): Run {
  return field(" AUTHOR ", cachedValue);
}

/** Create a TITLE field. */
export function titleField(cachedValue?: string): Run {
  return field(" TITLE ", cachedValue);
}

/** Create a SUBJECT field. */
export function subjectField(cachedValue?: string): Run {
  return field(" SUBJECT ", cachedValue);
}

/** Create a KEYWORDS field. */
export function keywordsField(cachedValue?: string): Run {
  return field(" KEYWORDS ", cachedValue);
}

/** Create a FILENAME field. */
export function fileNameField(options?: { includePath?: boolean; cachedValue?: string }): Run {
  const instruction = options?.includePath ? " FILENAME \\p " : " FILENAME ";
  return field(instruction, options?.cachedValue);
}

/** Create a FILESIZE field. */
export function fileSizeField(cachedValue?: string): Run {
  return field(" FILESIZE ", cachedValue);
}

/**
 * Create a STYLEREF field (references text from nearest paragraph with given style).
 *
 * Commonly used in headers to show current chapter/section heading.
 */
export function styleRefField(
  styleName: string,
  options?: {
    /** Search from bottom of page (default: top). */
    fromBottom?: boolean;
    /** Insert paragraph number instead of text. */
    insertParagraphNumber?: boolean;
    /** Insert position (above/below) reference. */
    insertPosition?: boolean;
    /** Suppress non-delimiter/non-numerical text. */
    suppressNonNumeric?: boolean;
    cachedValue?: string;
  }
): Run {
  let instruction = ` STYLEREF "${styleName}" `;
  if (options?.fromBottom) {
    instruction += "\\l ";
  }
  if (options?.insertParagraphNumber) {
    instruction += "\\n ";
  }
  if (options?.insertPosition) {
    instruction += "\\p ";
  }
  if (options?.suppressNonNumeric) {
    instruction += "\\t ";
  }
  return field(instruction, options?.cachedValue);
}

/**
 * Create a REF field (references a bookmark).
 */
export function refField(
  bookmarkName: string,
  options?: {
    /** Insert bookmark number. */
    insertNumber?: boolean;
    /** Hyperlink to bookmark. */
    hyperlink?: boolean;
    /** Paragraph number (no context). */
    paragraphNumber?: boolean;
    /** Relative paragraph number. */
    relativeParagraphNumber?: boolean;
    /** Full context paragraph number. */
    fullContext?: boolean;
    /** Suppress non-delimiter. */
    suppressNonDelimiter?: boolean;
    cachedValue?: string;
  }
): Run {
  let instruction = ` REF ${bookmarkName} `;
  if (options?.insertNumber) {
    instruction += "\\n ";
  }
  if (options?.hyperlink) {
    instruction += "\\h ";
  }
  if (options?.paragraphNumber) {
    instruction += "\\w ";
  }
  if (options?.relativeParagraphNumber) {
    instruction += "\\r ";
  }
  if (options?.fullContext) {
    instruction += "\\w ";
  }
  if (options?.suppressNonDelimiter) {
    instruction += "\\t ";
  }
  return field(instruction, options?.cachedValue);
}

/** Create a PAGEREF field (references page of bookmark). */
export function pageRefField(
  bookmarkName: string,
  options?: { hyperlink?: boolean; relativePosition?: boolean; cachedValue?: string }
): Run {
  let instruction = ` PAGEREF ${bookmarkName} `;
  if (options?.hyperlink) {
    instruction += "\\h ";
  }
  if (options?.relativePosition) {
    instruction += "\\p ";
  }
  return field(instruction, options?.cachedValue);
}

/** Create a NOTEREF field (references a footnote/endnote). */
export function noteRefField(
  bookmarkName: string,
  options?: { hyperlink?: boolean; insertNumberFormat?: boolean; cachedValue?: string }
): Run {
  let instruction = ` NOTEREF ${bookmarkName} `;
  if (options?.hyperlink) {
    instruction += "\\h ";
  }
  if (options?.insertNumberFormat) {
    instruction += "\\f ";
  }
  return field(instruction, options?.cachedValue);
}

/** Create a HYPERLINK field (alternative to w:hyperlink element). */
export function hyperlinkField(
  target: string,
  options?: {
    /** Anchor within target (bookmark). */
    anchor?: string;
    /** Open in new window. */
    newWindow?: boolean;
    /** Display text (if omitted, uses target). */
    displayText?: string;
    /** Screen tip tooltip. */
    tooltip?: string;
    cachedValue?: string;
  }
): Run {
  let instruction = ` HYPERLINK "${target}" `;
  if (options?.anchor) {
    instruction += `\\l "${options.anchor}" `;
  }
  if (options?.newWindow) {
    instruction += "\\n ";
  }
  if (options?.tooltip) {
    instruction += `\\o "${options.tooltip}" `;
  }
  return field(instruction, options?.cachedValue ?? options?.displayText);
}

/** Create a QUOTE field (literal text). */
export function quoteField(text: string, cachedValue?: string): Run {
  return field(` QUOTE "${text}" `, cachedValue ?? text);
}

/**
 * Create a TOC field for table of contents.
 *
 * @param options - TOC options.
 */
export function tocField(options?: {
  /** Heading levels to include (e.g. "1-3"). */
  headingLevels?: string;
  /** Include paragraphs with these styles. */
  styles?: readonly string[];
  /** Include table entries (TC fields) with these levels. */
  tcLevels?: string;
  /** Hyperlinks for entries. */
  hyperlink?: boolean;
  /** Right-align page numbers. */
  rightAlignedPageNumbers?: boolean;
  /** Use tab leader (default "." dots). */
  tabLeader?: "." | "-" | "_" | " ";
  /** Suppress page numbers. */
  noPageNumbers?: boolean;
  /** Identifier for table of captions. */
  captionIdentifier?: string;
  cachedValue?: string;
}): Run {
  let instruction = " TOC ";
  if (options?.headingLevels) {
    instruction += `\\o "${options.headingLevels}" `;
  }
  if (options?.styles && options.styles.length > 0) {
    instruction += `\\t "${options.styles.join(";")}" `;
  }
  if (options?.tcLevels) {
    instruction += `\\f ${options.tcLevels} `;
  }
  if (options?.hyperlink) {
    instruction += "\\h ";
  }
  if (options?.rightAlignedPageNumbers) {
    instruction += "\\z ";
  }
  if (options?.tabLeader) {
    instruction += `\\p "${options.tabLeader}" `;
  }
  if (options?.noPageNumbers) {
    instruction += "\\n ";
  }
  if (options?.captionIdentifier) {
    instruction += `\\c "${options.captionIdentifier}" `;
  }
  return field(instruction, options?.cachedValue);
}

/** Create a TC (table of contents entry) field. */
export function tcField(
  text: string,
  options?: { level?: number; suppressPageNumber?: boolean; cachedValue?: string }
): Run {
  let instruction = ` TC "${text}" `;
  if (options?.level !== undefined) {
    instruction += `\\l ${options.level} `;
  }
  if (options?.suppressPageNumber) {
    instruction += "\\n ";
  }
  return field(instruction, options?.cachedValue);
}

/** Create an XE (index entry) field. */
export function indexEntryField(
  text: string,
  options?: { bold?: boolean; italic?: boolean; cachedValue?: string }
): Run {
  let instruction = ` XE "${text}" `;
  if (options?.bold) {
    instruction += "\\b ";
  }
  if (options?.italic) {
    instruction += "\\i ";
  }
  return field(instruction, options?.cachedValue);
}

/** Create an INDEX field (renders index from XE entries). */
export function indexField(options?: {
  bookmark?: string;
  columns?: number;
  entryType?: string;
  cachedValue?: string;
}): Run {
  let instruction = " INDEX ";
  if (options?.bookmark) {
    instruction += `\\b ${options.bookmark} `;
  }
  if (options?.columns) {
    instruction += `\\c ${options.columns} `;
  }
  if (options?.entryType) {
    instruction += `\\f "${options.entryType}" `;
  }
  return field(instruction, options?.cachedValue);
}

/** Create an IF field (conditional content). */
export function ifField(
  condition: string,
  trueText: string,
  falseText: string,
  cachedValue?: string
): Run {
  return field(` IF ${condition} "${trueText}" "${falseText}" `, cachedValue);
}

/** Create an INCLUDETEXT field (includes external file content). */
export function includeTextField(
  filePath: string,
  options?: { bookmark?: string; cachedValue?: string }
): Run {
  let instruction = ` INCLUDETEXT "${filePath}" `;
  if (options?.bookmark) {
    instruction += `${options.bookmark} `;
  }
  return field(instruction, options?.cachedValue);
}

/** Create an INCLUDEPICTURE field. */
export function includePictureField(filePath: string, cachedValue?: string): Run {
  return field(` INCLUDEPICTURE "${filePath}" `, cachedValue);
}

/** Create a FORMTEXT field (legacy text form field). */
export function formTextField(cachedValue?: string): Run {
  return field(" FORMTEXT ", cachedValue);
}

/** Create a FORMCHECKBOX field (legacy checkbox form field). */
export function formCheckboxField(cachedValue?: string): Run {
  return field(" FORMCHECKBOX ", cachedValue);
}

/** Create a FORMDROPDOWN field (legacy dropdown form field). */
export function formDropdownField(cachedValue?: string): Run {
  return field(" FORMDROPDOWN ", cachedValue);
}

/** Create a paragraph. */
export function paragraph(children: ParagraphChild[], properties?: ParagraphProperties): Paragraph {
  return { type: "paragraph", properties, children };
}

/** Create a simple text paragraph. */
export function textParagraph(
  content: string,
  properties?: ParagraphProperties & { run?: RunProperties }
): Paragraph {
  const { run: runProps, ...pProps } = properties ?? {};
  return paragraph([text(content, runProps)], Object.keys(pProps).length > 0 ? pProps : undefined);
}

/** Create a heading paragraph. */
export function heading(content: string, level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9): Paragraph {
  return paragraph([text(content)], { style: `Heading${level}` });
}

/** Create a hyperlink. */
export function hyperlink(
  linkText: string,
  options: {
    rId?: string;
    url?: string;
    anchor?: string;
    tooltip?: string;
    properties?: RunProperties;
  }
): ParagraphChild {
  return {
    type: "hyperlink",
    rId: options.rId,
    url: options.url,
    anchor: options.anchor,
    tooltip: options.tooltip,
    children: [text(linkText, options.properties ?? { color: "0563C1", underline: "single" })]
  };
}

/** Create a bookmark start. */
export function bookmarkStart(id: number, name: string): ParagraphChild {
  return { type: "bookmarkStart", id, name };
}

/** Create a bookmark end. */
export function bookmarkEnd(id: number): ParagraphChild {
  return { type: "bookmarkEnd", id };
}

/** Create a comment range start marker. */
export function commentRangeStart(id: number): CommentRangeStart {
  return { type: "commentRangeStart", id };
}

/** Create a comment range end marker. */
export function commentRangeEnd(id: number): CommentRangeEnd {
  return { type: "commentRangeEnd", id };
}

/** Create a comment reference (inside paragraph children). */
export function commentReference(id: number): CommentReference {
  return { type: "commentReference", id };
}

/** Create an inserted run (track changes). */
export function insertedRun(run: Run, revision: RevisionInfo): InsertedRun {
  return { type: "insertedRun", revision, run };
}

/** Create a deleted run (track changes). */
export function deletedRun(run: Run, revision: RevisionInfo): DeletedRun {
  return { type: "deletedRun", revision, run };
}

/** Create a moved-from run (track changes — source of a move). */
export function movedFromRun(run: Run, revision: RevisionInfo): MovedFromRun {
  return { type: "movedFromRun", revision, run };
}

/** Create a moved-to run (track changes — destination of a move). */
export function movedToRun(run: Run, revision: RevisionInfo): MovedToRun {
  return { type: "movedToRun", revision, run };
}

/** Create a move range start marker. */
export function moveFromRangeStart(
  id: number,
  author: string,
  options?: { date?: string; name?: string }
): MoveRangeMarker {
  return { type: "moveFromRangeStart", id, author, date: options?.date, name: options?.name };
}

/** Create a move range end marker. */
export function moveFromRangeEnd(id: number): MoveRangeMarker {
  return { type: "moveFromRangeEnd", id };
}

/** Create a move-to range start marker. */
export function moveToRangeStart(
  id: number,
  author: string,
  options?: { date?: string; name?: string }
): MoveRangeMarker {
  return { type: "moveToRangeStart", id, author, date: options?.date, name: options?.name };
}

/** Create a move-to range end marker. */
export function moveToRangeEnd(id: number): MoveRangeMarker {
  return { type: "moveToRangeEnd", id };
}

/** Create a checkbox. */
export function checkBox(options?: {
  checked?: boolean;
  checkedState?: { value: string; font?: string };
  uncheckedState?: { value: string; font?: string };
}): CheckBox {
  return {
    type: "checkBox",
    checked: options?.checked,
    checkedState: options?.checkedState,
    uncheckedState: options?.uncheckedState
  };
}

// =============================================================================
// Math Builders
// =============================================================================

/** Create a math block. */
export function mathBlock(content: MathContent[]): MathBlock {
  return { type: "math", content };
}

/** Create a math text run. */
export function mathRun(
  mathText: string,
  properties?: { italic?: boolean; bold?: boolean; font?: string }
): MathContent {
  return { type: "mathRun", text: mathText, properties };
}

/** Create a math fraction. */
export function mathFraction(
  numerator: MathContent[],
  denominator: MathContent[],
  fractionType?: "bar" | "skw" | "lin" | "noBar"
): MathContent {
  return { type: "mathFraction", fractionType, numerator, denominator };
}

/** Create a math square root. */
export function mathSqrt(content: MathContent[]): MathContent {
  return { type: "mathRadical", content, hideDegree: true };
}

/** Create a math nth root. */
export function mathRoot(degree: MathContent[], content: MathContent[]): MathContent {
  return { type: "mathRadical", degree, content };
}

/** Create a math summation. */
export function mathSum(
  content: MathContent[],
  sub?: MathContent[],
  sup?: MathContent[]
): MathContent {
  return { type: "mathNary", char: "\u2211", sub, sup, content };
}

/** Create a math integral. */
export function mathIntegral(
  content: MathContent[],
  sub?: MathContent[],
  sup?: MathContent[]
): MathContent {
  return { type: "mathNary", char: "\u222B", sub, sup, content };
}

/** Create a math product. */
export function mathProduct(
  content: MathContent[],
  sub?: MathContent[],
  sup?: MathContent[]
): MathContent {
  return { type: "mathNary", char: "\u220F", sub, sup, content };
}

/** Create a math superscript. */
export function mathSuperScript(base: MathContent[], superScript: MathContent[]): MathContent {
  return { type: "mathSuperScript", base, superScript };
}

/** Create a math subscript. */
export function mathSubScript(base: MathContent[], subScript: MathContent[]): MathContent {
  return { type: "mathSubScript", base, subScript };
}

/** Create a math sub-superscript. */
export function mathSubSuperScript(
  base: MathContent[],
  subScript: MathContent[],
  superScript: MathContent[]
): MathContent {
  return { type: "mathSubSuperScript", base, subScript, superScript };
}

/** Create a math pre-sub-superscript (subscript/superscript before the base). */
export function mathPreSubSuperScript(
  base: MathContent[],
  preSubScript: MathContent[],
  preSuperScript: MathContent[]
): MathContent {
  return { type: "mathPreSubSuperScript", base, preSubScript, preSuperScript };
}

/** Create a math phantom (invisible expression that takes up space). */
export function mathPhantom(
  content: MathContent[],
  options?: {
    show?: boolean;
    zeroWidth?: boolean;
    zeroAscent?: boolean;
    zeroDescent?: boolean;
    transparent?: boolean;
  }
): MathContent {
  return { type: "mathPhantom", content, ...options };
}

/** Create a math group character (e.g. a horizontal brace over an expression). */
export function mathGroupChar(
  base: MathContent[],
  options?: {
    char?: string;
    position?: "top" | "bottom";
    verticalAlign?: "top" | "center" | "bottom";
  }
): MathContent {
  return { type: "mathGroupChar", base, ...options };
}

/** Create a math border box (draw borders around / strike through an expression). */
export function mathBorderBox(
  content: MathContent[],
  options?: {
    hideTop?: boolean;
    hideBottom?: boolean;
    hideLeft?: boolean;
    hideRight?: boolean;
    strikeBlTr?: boolean;
    strikeTlBr?: boolean;
    strikeH?: boolean;
    strikeV?: boolean;
  }
): MathContent {
  return { type: "mathBorderBox", content, ...options };
}

/** Create a math delimiter (parentheses, brackets, etc.). */
export function mathDelimiter(
  content: MathContent[][],
  options?: { beginChar?: string; endChar?: string; separatorChar?: string }
): MathContent {
  return {
    type: "mathDelimiter",
    beginChar: options?.beginChar,
    endChar: options?.endChar,
    separatorChar: options?.separatorChar,
    content
  };
}

/** Create a math n-ary operator (sum, integral, product, etc.). */
export function mathNary(
  char: string,
  content: MathContent[],
  sub?: MathContent[],
  sup?: MathContent[]
): MathContent {
  return { type: "mathNary", char, sub, sup, content };
}

/** Create a math function (sin, cos, lim, etc.). */
export function mathFunction(name: MathContent[], content: MathContent[]): MathContent {
  return { type: "mathFunction", name, content };
}

/** Create a math limit (upper or lower). */
export function mathLimit(
  base: MathContent[],
  limit: MathContent[],
  limitType: "upper" | "lower" = "lower"
): MathContent {
  return { type: "mathLimit", base, limit, limitType };
}

/** Create a math matrix. */
export function mathMatrix(rows: MathContent[][][]): MathContent {
  return { type: "mathMatrix", rows };
}

/** Create a math accent (hat, tilde, etc.). */
export function mathAccent(content: MathContent[], char?: string): MathContent {
  return { type: "mathAccent", content, char };
}

/** Create a math bar (overbar/underbar). */
export function mathBar(content: MathContent[], position?: "top" | "bottom"): MathContent {
  return { type: "mathBar", content, position: position ?? "top" };
}

/** Create a math box. */
export function mathBox(content: MathContent[]): MathContent {
  return { type: "mathBox", content };
}

/** Create a math equation array. */
export function mathEquationArray(rows: MathContent[][]): MathContent {
  return { type: "mathEquationArray", rows };
}

/** Create a symbol run. */
export function symbol(font: string, char: string, properties?: RunProperties): Run {
  return { properties, content: [{ type: "symbol", font, char }] };
}

/** Create a floating image (body-level). */
export function floatingImage(options: {
  rId: string;
  width: Emu;
  height: Emu;
  horizontalPosition?: FloatingImage["horizontalPosition"];
  verticalPosition?: FloatingImage["verticalPosition"];
  wrap?: FloatingImage["wrap"];
  altText?: string;
  name?: string;
  behindDoc?: boolean;
  lockAnchor?: boolean;
  layoutInCell?: boolean;
  allowOverlap?: boolean;
  simplePos?: { x: Emu; y: Emu };
  distT?: Emu;
  distB?: Emu;
  distL?: Emu;
  distR?: Emu;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  srcRect?: FloatingImage["srcRect"];
}): FloatingImage {
  return {
    type: "floatingImage",
    rId: options.rId,
    width: options.width,
    height: options.height,
    horizontalPosition: options.horizontalPosition ?? { relativeTo: "column", offset: 0 },
    verticalPosition: options.verticalPosition ?? { relativeTo: "paragraph", offset: 0 },
    wrap: options.wrap ?? { style: "square" },
    altText: options.altText,
    name: options.name,
    behindDoc: options.behindDoc,
    lockAnchor: options.lockAnchor,
    layoutInCell: options.layoutInCell,
    allowOverlap: options.allowOverlap,
    simplePos: options.simplePos,
    distT: options.distT,
    distB: options.distB,
    distL: options.distL,
    distR: options.distR,
    rotation: options.rotation,
    flipHorizontal: options.flipHorizontal,
    flipVertical: options.flipVertical,
    srcRect: options.srcRect
  };
}

/** Create a DrawingML shape. */
export function drawingShape(options: {
  shapeType: DrawingShape["shapeType"];
  width: Emu;
  height: Emu;
  fillColor?: HexColor;
  noFill?: boolean;
  outlineColor?: HexColor;
  outlineWidth?: Emu;
  noOutline?: boolean;
  textContent?: readonly Paragraph[];
  altText?: string;
  name?: string;
  horizontalPosition?: DrawingShape["horizontalPosition"];
  verticalPosition?: DrawingShape["verticalPosition"];
  wrap?: DrawingShape["wrap"];
  behindDoc?: boolean;
  rotation?: number;
}): DrawingShape {
  return {
    type: "drawingShape",
    shapeType: options.shapeType,
    width: options.width,
    height: options.height,
    fillColor: options.fillColor,
    noFill: options.noFill,
    outlineColor: options.outlineColor,
    outlineWidth: options.outlineWidth,
    noOutline: options.noOutline,
    textContent: options.textContent,
    altText: options.altText,
    name: options.name,
    horizontalPosition: options.horizontalPosition ?? { relativeTo: "column", offset: 0 },
    verticalPosition: options.verticalPosition ?? { relativeTo: "paragraph", offset: 0 },
    wrap: options.wrap ?? { style: "square" },
    behindDoc: options.behindDoc,
    rotation: options.rotation
  };
}

/** Create a chart content block. */
export function chart(options: {
  type: Chart["type"];
  series: readonly ChartSeries[];
  title?: string;
  legend?: ChartLegendPosition;
  categoryAxis?: ChartAxis;
  valueAxis?: ChartAxis;
  plotAreaColor?: HexColor;
  chartAreaColor?: HexColor;
  view3d?: boolean;
  style?: number;
  width?: Emu;
  height?: Emu;
  altText?: string;
  name?: string;
}): ChartContent {
  return {
    type: "chart",
    chart: {
      type: options.type,
      series: options.series,
      title: options.title,
      legend: options.legend,
      categoryAxis: options.categoryAxis,
      valueAxis: options.valueAxis,
      plotAreaColor: options.plotAreaColor,
      chartAreaColor: options.chartAreaColor,
      view3d: options.view3d,
      style: options.style,
      width: options.width,
      height: options.height
    },
    altText: options.altText,
    name: options.name
  };
}

/** Create a structured document tag (content control). */
export function structuredDocumentTag(
  content: (Paragraph | Table)[],
  properties?: SdtProperties
): StructuredDocumentTag {
  return { type: "sdt", properties: properties ?? {}, content };
}

// =============================================================================
// Table Builders
// =============================================================================

/** Shorthand border. */
export function border(style: Border["style"] = "single", size = 4, color = "auto"): Border {
  return { style, size, space: 0, color };
}

/** Create standard grid borders for a table. */
export function gridBorders(size = 4, color = "auto"): TableBorders {
  const b = border("single", size, color);
  return { top: b, left: b, bottom: b, right: b, insideH: b, insideV: b };
}

/** Create a table cell. */
export function cell(
  content: string | (Paragraph | Table)[],
  properties?: TableCellProperties
): TableCell {
  if (typeof content === "string") {
    return { properties, content: [textParagraph(content)] };
  }
  return { properties, content };
}

/** Create a table row. */
export function row(cells: TableCell[], properties?: TableRowProperties): TableRow {
  return { properties, cells };
}

/** Create a table. */
export function table(
  rows: TableRow[],
  properties?: TableProperties,
  columnWidths?: Twips[]
): Table {
  return { type: "table", properties, columnWidths, rows };
}

/** Create a simple table from a 2D string array. */
export function simpleTable(
  data: string[][],
  options?: {
    headerRow?: boolean;
    borders?: boolean;
    width?: TableWidth;
    columnWidths?: Twips[];
  }
): Table {
  const opts = { headerRow: true, borders: true, ...options };
  const tableRows: TableRow[] = data.map((rowData, rowIndex) => {
    const cells = rowData.map(cellText => cell(cellText));
    return row(cells, rowIndex === 0 && opts.headerRow ? { tableHeader: true } : undefined);
  });

  return table(
    tableRows,
    {
      width: opts.width ?? { value: 5000, type: "pct" },
      borders: opts.borders ? gridBorders() : undefined
    },
    opts.columnWidths
  );
}

// =============================================================================
// Document Builder Class
// =============================================================================

/**
 * Fluent builder for constructing DOCX documents.
 *
 * @example
 * ```ts
 * const doc = new DocumentBuilder()
 *   .addHeading("Hello World", 1)
 *   .addParagraph("This is a paragraph.")
 *   .addTable([["Name", "Age"], ["Alice", "30"]])
 *   .build();
 *
 * const bytes = await doc.toBuffer();
 * ```
 */
export class DocumentBuilder {
  private _body: BodyContent[] = [];
  private _sectionProperties?: SectionProperties;
  private _styles: StyleDef[] = [];
  private _docDefaults?: DocDefaults;
  private _abstractNumberings: AbstractNumbering[] = [];
  private _numberingInstances: NumberingInstance[] = [];
  private _headers = new Map<string, HeaderDef>();
  private _footers = new Map<string, FooterDef>();
  private _footnotes: FootnoteDef[] = [];
  private _endnotes: EndnoteDef[] = [];
  private _images: ImageDef[] = [];
  private _fonts: FontDef[] = [];
  private _settings?: DocumentSettings;
  private _coreProperties?: CoreProperties;
  private _appProperties?: AppProperties;
  private _comments: CommentDef[] = [];
  private _background?: DocumentBackground;
  private _customProperties: CustomProperty[] = [];
  private _watermark?: Watermark;
  private _nextImageId = 1;
  private _nextFootnoteId = 1;
  private _nextEndnoteId = 1;
  private _nextBookmarkId = 0;
  private _nextAbstractNumId = 0;
  private _nextNumId = 1;
  private _nextDrawingId = 1;
  private _nextCommentId = 0;

  /** Add raw body content. */
  addContent(content: BodyContent): this {
    this._body.push(content);
    return this;
  }

  /** Add a paragraph with runs. */
  addParagraphElement(para: Paragraph): this {
    this._body.push(para);
    return this;
  }

  /** Add a simple text paragraph. */
  addParagraph(content: string, properties?: ParagraphProperties & { run?: RunProperties }): this {
    this._body.push(textParagraph(content, properties));
    return this;
  }

  /** Add a heading. */
  addHeading(content: string, level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 = 1): this {
    this._body.push(heading(content, level));
    return this;
  }

  /** Add a page break. */
  addPageBreak(): this {
    this._body.push(paragraph([pageBreak()]));
    return this;
  }

  /** Add a table from a 2D array. */
  addTable(
    data: string[][],
    options?: { headerRow?: boolean; borders?: boolean; width?: TableWidth; columnWidths?: Twips[] }
  ): this {
    this._body.push(simpleTable(data, options));
    return this;
  }

  /** Add a table element. */
  addTableElement(tbl: Table): this {
    this._body.push(tbl);
    return this;
  }

  /** Add an inline image. Returns the image relationship ID and drawing ID. */
  addImage(
    data: Uint8Array,
    mediaType: ImageMediaType,
    width: Emu,
    height: Emu,
    options?: { altText?: string; name?: string }
  ): { rId: string; drawingId: number } {
    const fileName = `image${this._nextImageId}.${mediaType}`;
    const rId = `__img_${this._nextImageId}`;
    const drawingId = this._nextDrawingId++;

    this._images.push({ data, mediaType, fileName, rId });

    this._body.push(
      paragraph([
        {
          content: [
            {
              type: "image",
              rId,
              width,
              height,
              altText: options?.altText,
              name: options?.name ?? `Picture ${this._nextImageId}`,
              drawingId
            }
          ]
        }
      ])
    );

    this._nextImageId++;
    return { rId, drawingId };
  }

  /** Add a floating image. Returns the image relationship ID. */
  addFloatingImage(
    data: Uint8Array,
    mediaType: ImageMediaType,
    width: Emu,
    height: Emu,
    options?: {
      altText?: string;
      name?: string;
      horizontalPosition?: FloatingImage["horizontalPosition"];
      verticalPosition?: FloatingImage["verticalPosition"];
      wrap?: FloatingImage["wrap"];
      behindDoc?: boolean;
      lockAnchor?: boolean;
      layoutInCell?: boolean;
      allowOverlap?: boolean;
      distT?: Emu;
      distB?: Emu;
      distL?: Emu;
      distR?: Emu;
      rotation?: number;
      flipHorizontal?: boolean;
      flipVertical?: boolean;
    }
  ): string {
    const fileName = `image${this._nextImageId}.${mediaType}`;
    const rId = `__img_${this._nextImageId}`;

    this._images.push({ data, mediaType, fileName, rId });

    this._body.push(
      floatingImage({
        rId,
        width,
        height,
        altText: options?.altText,
        name: options?.name ?? `Picture ${this._nextImageId}`,
        horizontalPosition: options?.horizontalPosition,
        verticalPosition: options?.verticalPosition,
        wrap: options?.wrap,
        behindDoc: options?.behindDoc,
        lockAnchor: options?.lockAnchor,
        layoutInCell: options?.layoutInCell,
        allowOverlap: options?.allowOverlap,
        distT: options?.distT,
        distB: options?.distB,
        distL: options?.distL,
        distR: options?.distR,
        rotation: options?.rotation,
        flipHorizontal: options?.flipHorizontal,
        flipVertical: options?.flipVertical
      })
    );

    this._nextImageId++;
    return rId;
  }

  /** Add a custom font definition. */
  addFont(font: FontDef): this {
    this._fonts.push(font);
    return this;
  }

  /** Set a text watermark on the document. */
  setWatermark(watermark: Watermark): this {
    this._watermark = watermark;
    return this;
  }

  /** Add a footnote. Returns the footnote ID. */
  addFootnote(content: string | Paragraph[]): number {
    const id = this._nextFootnoteId++;
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    this._footnotes.push({ id, content: paras });
    return id;
  }

  /** Add an endnote. Returns the endnote ID. */
  addEndnote(content: string | Paragraph[]): number {
    const id = this._nextEndnoteId++;
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    this._endnotes.push({ id, content: paras });
    return id;
  }

  /** Add a comment. Returns the comment ID. */
  addComment(
    author: string,
    content: string | Paragraph[],
    options?: { date?: string; initials?: string }
  ): number {
    const id = this._nextCommentId++;
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    this._comments.push({
      id,
      author,
      date: options?.date,
      initials: options?.initials,
      content: paras
    });
    return id;
  }

  /** Add a Table of Contents. */
  addTableOfContents(options?: Partial<Omit<TableOfContents, "type">>): this {
    this._body.push({
      type: "tableOfContents",
      headingStyleRange: options?.headingStyleRange ?? "1-3",
      hyperlink: options?.hyperlink ?? true,
      ...options
    });
    return this;
  }

  /** Add a math equation block. */
  addMath(content: MathContent[]): this {
    this._body.push(mathBlock(content));
    return this;
  }

  /** Add a text box. */
  addTextBox(
    content: string | Paragraph[],
    options?: { width?: Twips; height?: Twips; stroke?: boolean; fill?: boolean }
  ): this {
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    this._body.push({
      type: "textBox",
      content: paras,
      width: options?.width,
      height: options?.height,
      stroke: options?.stroke,
      fill: options?.fill
    });
    return this;
  }

  /** Add a bullet list. */
  addBulletList(items: string[], level = 0): this {
    // Create abstract numbering for bullets if not exists
    let bulletAbsId = this._abstractNumberings.find(
      a => a.levels[0]?.format === "bullet"
    )?.abstractNumId;

    if (bulletAbsId === undefined) {
      bulletAbsId = this._nextAbstractNumId++;
      this._abstractNumberings.push({
        abstractNumId: bulletAbsId,
        multiLevelType: "hybridMultilevel",
        levels: [
          {
            level: 0,
            start: 1,
            format: "bullet",
            text: "\uF0B7",
            justification: "left",
            paragraphProperties: { indent: { left: 720, hanging: 360 } },
            runProperties: { font: { ascii: "Symbol", hAnsi: "Symbol" } }
          },
          {
            level: 1,
            start: 1,
            format: "bullet",
            text: "o",
            justification: "left",
            paragraphProperties: { indent: { left: 1440, hanging: 360 } },
            runProperties: { font: { ascii: "Courier New", hAnsi: "Courier New" } }
          },
          {
            level: 2,
            start: 1,
            format: "bullet",
            text: "\uF0A7",
            justification: "left",
            paragraphProperties: { indent: { left: 2160, hanging: 360 } },
            runProperties: { font: { ascii: "Wingdings", hAnsi: "Wingdings" } }
          }
        ]
      });
      this._numberingInstances.push({
        numId: this._nextNumId++,
        abstractNumId: bulletAbsId
      });
    }

    const numId = this._numberingInstances.find(n => n.abstractNumId === bulletAbsId)!.numId;

    for (const item of items) {
      this._body.push(textParagraph(item, { numbering: { numId, level } }));
    }

    return this;
  }

  /** Add a numbered list. */
  addNumberedList(items: string[], level = 0): this {
    let numAbsId = this._abstractNumberings.find(
      a => a.levels[0]?.format === "decimal"
    )?.abstractNumId;

    if (numAbsId === undefined) {
      numAbsId = this._nextAbstractNumId++;
      this._abstractNumberings.push({
        abstractNumId: numAbsId,
        multiLevelType: "hybridMultilevel",
        levels: [
          {
            level: 0,
            start: 1,
            format: "decimal",
            text: "%1.",
            justification: "left",
            paragraphProperties: { indent: { left: 720, hanging: 360 } }
          },
          {
            level: 1,
            start: 1,
            format: "lowerLetter",
            text: "%2.",
            justification: "left",
            paragraphProperties: { indent: { left: 1440, hanging: 360 } }
          },
          {
            level: 2,
            start: 1,
            format: "lowerRoman",
            text: "%3.",
            justification: "right",
            paragraphProperties: { indent: { left: 2160, hanging: 180 } }
          }
        ]
      });
      this._numberingInstances.push({
        numId: this._nextNumId++,
        abstractNumId: numAbsId
      });
    }

    const numId = this._numberingInstances.find(n => n.abstractNumId === numAbsId)!.numId;

    for (const item of items) {
      this._body.push(textParagraph(item, { numbering: { numId, level } }));
    }

    return this;
  }

  /** Set section properties (page size, margins, etc.). */
  setSectionProperties(props: SectionProperties): this {
    this._sectionProperties = props;
    return this;
  }

  /** Set document defaults. */
  setDocDefaults(defaults: DocDefaults): this {
    this._docDefaults = defaults;
    return this;
  }

  /** Add a style definition. */
  addStyle(style: StyleDef): this {
    this._styles.push(style);
    return this;
  }

  /** Set default styles (Normal, Heading1-6, Hyperlink, etc.). */
  useDefaultStyles(): this {
    this._docDefaults = {
      runProperties: {
        font: { ascii: "Calibri", hAnsi: "Calibri", eastAsia: "SimSun", cs: "Times New Roman" },
        size: 22,
        sizeCs: 22,
        language: { val: "en-US" }
      },
      paragraphProperties: {
        spacing: { after: 160, line: 259, lineRule: "auto" }
      }
    };

    this._styles.push(
      { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true, qFormat: true },
      {
        type: "paragraph",
        styleId: "Heading1",
        name: "heading 1",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        paragraphProperties: {
          keepNext: true,
          keepLines: true,
          spacing: { before: 240, after: 0 }
        },
        runProperties: { font: "Calibri Light", color: "2F5496", size: 32 }
      },
      {
        type: "paragraph",
        styleId: "Heading2",
        name: "heading 2",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        paragraphProperties: { keepNext: true, keepLines: true, spacing: { before: 40, after: 0 } },
        runProperties: { font: "Calibri Light", color: "2F5496", size: 26 }
      },
      {
        type: "paragraph",
        styleId: "Heading3",
        name: "heading 3",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        paragraphProperties: { keepNext: true, keepLines: true, spacing: { before: 40, after: 0 } },
        runProperties: { font: "Calibri Light", color: "1F3763", size: 24 }
      },
      {
        type: "character",
        styleId: "Hyperlink",
        name: "Hyperlink",
        uiPriority: 99,
        runProperties: { color: "0563C1", underline: "single" }
      },
      {
        type: "table",
        styleId: "TableGrid",
        name: "Table Grid",
        basedOn: "TableNormal",
        uiPriority: 39,
        tableProperties: { borders: gridBorders(4, "auto") }
      }
    );

    return this;
  }

  /** Set a header for the given type. */
  setHeader(type: string, content: HeaderFooterContent): this {
    this._headers.set(type, { content });
    return this;
  }

  /** Set a footer for the given type. */
  setFooter(type: string, content: HeaderFooterContent): this {
    this._footers.set(type, { content });
    return this;
  }

  /** Set document settings. */
  setSettings(settings: DocumentSettings): this {
    this._settings = settings;
    return this;
  }

  /** Set core properties (metadata). */
  setCoreProperties(props: CoreProperties): this {
    this._coreProperties = props;
    return this;
  }

  /** Set application properties. */
  setAppProperties(props: AppProperties): this {
    this._appProperties = props;
    return this;
  }

  /** Set document background. */
  setBackground(background: DocumentBackground): this {
    this._background = background;
    return this;
  }

  /** Add a custom document property. */
  addCustomProperty(name: string, value: CustomPropertyValue): this {
    this._customProperties.push({ name, value });
    return this;
  }

  /** Add a section break with properties. */
  addSectionBreak(props: SectionProperties): this {
    // Insert as the last paragraph's section properties
    if (this._body.length > 0) {
      const last = this._body[this._body.length - 1];
      if (last.type === "paragraph") {
        const existingProps = last.properties ?? {};
        (this._body[this._body.length - 1] as any) = {
          ...last,
          properties: { ...existingProps, sectionProperties: props }
        };
        return this;
      }
    }
    // If no previous paragraph, add an empty one with section properties
    this._body.push(paragraph([], { sectionProperties: props }));
    return this;
  }

  /** Get next available bookmark ID. */
  nextBookmarkId(): number {
    return this._nextBookmarkId++;
  }

  /** Build the DocxDocument model. */
  build(): DocxDocument {
    return {
      body: this._body,
      sectionProperties: this._sectionProperties ?? {
        pageSize: { width: 12240, height: 15840 },
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      },
      styles: this._styles.length > 0 ? this._styles : undefined,
      docDefaults: this._docDefaults,
      abstractNumberings:
        this._abstractNumberings.length > 0 ? this._abstractNumberings : undefined,
      numberingInstances:
        this._numberingInstances.length > 0 ? this._numberingInstances : undefined,
      headers: this._headers.size > 0 ? this._headers : undefined,
      footers: this._footers.size > 0 ? this._footers : undefined,
      footnotes: this._footnotes.length > 0 ? this._footnotes : undefined,
      endnotes: this._endnotes.length > 0 ? this._endnotes : undefined,
      images: this._images.length > 0 ? this._images : undefined,
      fonts: this._fonts.length > 0 ? this._fonts : undefined,
      settings: this._settings,
      coreProperties: this._coreProperties,
      appProperties: this._appProperties,
      comments: this._comments.length > 0 ? this._comments : undefined,
      background: this._background,
      customProperties: this._customProperties.length > 0 ? this._customProperties : undefined,
      watermark: this._watermark
    };
  }

  /** Build and package to DOCX bytes. */
  async toBuffer(compressionLevel?: number): Promise<Uint8Array> {
    return packageDocx(this.build(), compressionLevel);
  }

  /** Build and package to base64 string. */
  async toBase64(compressionLevel?: number): Promise<string> {
    const bytes = await this.toBuffer(compressionLevel);
    return bytesToBase64(bytes);
  }
}

// =============================================================================
// Theme Color Resolution
// =============================================================================

/**
 * Map OOXML theme color attribute names to theme color scheme keys.
 * Word uses different names in run/paragraph properties vs the theme XML.
 */
const THEME_COLOR_MAP: Record<string, string> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
  // Direct names also work
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
  hlink: "hlink",
  folHlink: "folHlink"
};

/**
 * Resolve a ColorSpec to an actual hex RGB color using the document theme.
 *
 * Applies theme color lookup + tint/shade transformations per OOXML spec.
 *
 * @param color - The color value (HexColor string or ColorSpec).
 * @param theme - The document theme (from `doc.theme`).
 * @returns Resolved hex color string (6 chars, no #), or undefined if unresolvable.
 */
export function resolveThemeColor(
  color: HexColor | ColorSpec | undefined,
  theme?: DocumentTheme
): HexColor | undefined {
  if (color === undefined) {
    return undefined;
  }
  if (typeof color === "string") {
    return color;
  }
  // ColorSpec with val — use directly
  if (color.val && color.val !== "auto") {
    return color.val;
  }
  // Resolve via theme
  if (!color.themeColor || !theme) {
    return color.val;
  }
  const key = THEME_COLOR_MAP[color.themeColor] ?? color.themeColor;
  const base = (theme.colorScheme.colors as Record<string, string>)[key];
  if (!base) {
    return color.val;
  }
  // Apply tint or shade
  if (color.themeTint) {
    return applyTint(base, parseInt(color.themeTint, 16) / 255);
  }
  if (color.themeShade) {
    return applyShade(base, parseInt(color.themeShade, 16) / 255);
  }
  return base;
}

/** Apply tint to a hex color. tint=1 → white, tint=0 → original. */
function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const nr = Math.round(r + (255 - r) * tint);
  const ng = Math.round(g + (255 - g) * tint);
  const nb = Math.round(b + (255 - b) * tint);
  return toHex2(nr) + toHex2(ng) + toHex2(nb);
}

/** Apply shade to a hex color. shade=1 → original, shade=0 → black. */
function applyShade(hex: string, shade: number): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const nr = Math.round(r * shade);
  const ng = Math.round(g * shade);
  const nb = Math.round(b * shade);
  return toHex2(nr) + toHex2(ng) + toHex2(nb);
}

function toHex2(n: number): string {
  const h = Math.max(0, Math.min(255, n)).toString(16);
  return h.length < 2 ? "0" + h : h;
}

// =============================================================================
// Search & Replace
// =============================================================================

/** Result of a text search in a document. */
export interface SearchResult {
  /** Paragraph index in body. */
  readonly paragraphIndex: number;
  /** The matched text. */
  readonly match: string;
  /** Character offset within the paragraph's concatenated text. */
  readonly offset: number;
}

/** Extract concatenated plain text from a paragraph's runs. */
function paragraphText(para: Paragraph): string {
  let text = "";
  for (const child of para.children) {
    if ("content" in child && Array.isArray(child.content)) {
      for (const c of child.content) {
        if ("type" in c && c.type === "text" && "text" in c) {
          text += (c as { text: string }).text;
        }
      }
    }
  }
  return text;
}

// =============================================================================
// Document Query API (read-only helpers)
// =============================================================================

/**
 * Count all top-level paragraphs in the document body.
 */
export function paragraphCount(doc: DocxDocument): number {
  let count = 0;
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      count++;
    }
  }
  return count;
}

/**
 * Count words across all paragraphs in the document body.
 * Uses simple whitespace splitting; for East Asian text, each CJK character
 * is counted as one "word" to approximate meaningful unit count.
 */
export function countWords(doc: DocxDocument): number {
  let count = 0;
  const cjkRe = /[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g;
  const walkParagraphs = (blocks: readonly BodyContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        const text = paragraphText(block);
        // Count CJK chars + latin word-like tokens
        const cjkCount = (text.match(cjkRe) ?? []).length;
        const latin = text.replace(cjkRe, " ").trim();
        const latinCount = latin ? latin.split(/\s+/).length : 0;
        count += cjkCount + latinCount;
      } else if (block.type === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walkParagraphs(cell.content as readonly BodyContent[]);
          }
        }
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table")
        );
        walkParagraphs(filtered as readonly BodyContent[]);
      }
    }
  };
  walkParagraphs(doc.body);
  return count;
}

/** A heading extracted from a document. */
export interface DocumentHeading {
  /** Heading level (1-9). */
  readonly level: number;
  /** Plain text of the heading. */
  readonly text: string;
  /** Index into doc.body where the paragraph resides. */
  readonly paragraphIndex: number;
  /** Style ID used (e.g. "Heading1"). */
  readonly style?: string;
}

/**
 * Extract the heading outline from a document.
 *
 * Matches paragraphs whose style is `Heading1` through `Heading9` (case-insensitive),
 * or whose `outlineLevel` property is set (0-8).
 */
export function getHeadings(doc: DocxDocument): DocumentHeading[] {
  const out: DocumentHeading[] = [];
  doc.body.forEach((block, i) => {
    if (block.type !== "paragraph") {
      return;
    }
    const style = block.properties?.style;
    const styleMatch = style ? /^Heading\s*(\d)$/i.exec(style) : null;
    let level: number | undefined;
    if (styleMatch) {
      level = parseInt(styleMatch[1], 10);
    } else if (block.properties?.outlineLevel !== undefined && block.properties.outlineLevel < 9) {
      level = block.properties.outlineLevel + 1;
    }
    if (level !== undefined && level >= 1 && level <= 9) {
      out.push({
        level,
        text: paragraphText(block),
        paragraphIndex: i,
        style
      });
    }
  });
  return out;
}

/**
 * Find a bookmark by name.
 *
 * @returns The bookmark start marker + its location, or `undefined` if not found.
 */
export function findBookmark(
  doc: DocxDocument,
  name: string
): { bookmark: BookmarkStart; paragraphIndex: number; childIndex: number } | undefined {
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    if (block.type !== "paragraph") {
      continue;
    }
    for (let j = 0; j < block.children.length; j++) {
      const ch = block.children[j];
      if ("type" in ch && ch.type === "bookmarkStart" && ch.name === name) {
        return { bookmark: ch, paragraphIndex: i, childIndex: j };
      }
    }
  }
  return undefined;
}

/**
 * Find a comment by its ID.
 */
export function findComment(doc: DocxDocument, id: number): CommentDef | undefined {
  return doc.comments?.find(c => c.id === id);
}

/**
 * List all images registered in the document.
 */
export function listImages(doc: DocxDocument): readonly ImageDef[] {
  return doc.images ?? [];
}

/**
 * List all tables in the document body (top-level only).
 */
export function listTables(doc: DocxDocument): readonly Table[] {
  return doc.body.filter((b): b is Table => b.type === "table");
}

/**
 * Collect all hyperlinks in the document body.
 */
export function listHyperlinks(doc: DocxDocument): readonly Hyperlink[] {
  const out: Hyperlink[] = [];
  const walk = (children: readonly ParagraphChild[]): void => {
    for (const ch of children) {
      if ("type" in ch && ch.type === "hyperlink") {
        out.push(ch);
        walk(ch.children as unknown as readonly ParagraphChild[]);
      }
    }
  };
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      walk(block.children);
    }
  }
  return out;
}

/**
 * Get the total number of tables (top-level) and nested tables.
 */
export function tableCount(doc: DocxDocument): number {
  let count = 0;
  const walk = (blocks: readonly BodyContent[]): void => {
    for (const block of blocks) {
      if (block.type === "table") {
        count++;
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walk(cell.content as readonly BodyContent[]);
          }
        }
      }
    }
  };
  walk(doc.body);
  return count;
}

/**
 * Extract plain text from the entire document body.
 *
 * Paragraphs are separated by `\n`. Tables render as tab-separated cell text.
 */
export function extractText(doc: DocxDocument): string {
  const lines: string[] = [];
  const walk = (blocks: readonly BodyContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        lines.push(paragraphText(block));
      } else if (block.type === "table") {
        for (const row of block.rows) {
          const cellTexts: string[] = [];
          for (const cell of row.cells) {
            const cellLines: string[] = [];
            walk(cell.content as readonly BodyContent[]);
            // collect last paragraphs text
            cellTexts.push(cellLines.join(" "));
          }
          lines.push(cellTexts.join("\t"));
        }
      }
    }
  };
  walk(doc.body);
  return lines.join("\n");
}

/**
 * Search for text occurrences in a document's body paragraphs.
 *
 * @param doc - The document model to search.
 * @param query - String or RegExp to search for.
 * @returns Array of search results.
 */
export function searchText(doc: DocxDocument, query: string | RegExp): SearchResult[] {
  const results: SearchResult[] = [];
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    if (block.type !== "paragraph") {
      continue;
    }
    const text = paragraphText(block);
    if (typeof query === "string") {
      let idx = text.indexOf(query);
      while (idx !== -1) {
        results.push({ paragraphIndex: i, match: query, offset: idx });
        idx = text.indexOf(query, idx + 1);
      }
    } else {
      const re = new RegExp(
        query.source,
        query.flags.includes("g") ? query.flags : query.flags + "g"
      );
      let m: RegExpExecArray | null;
      while ((m = re.exec(text)) !== null) {
        results.push({ paragraphIndex: i, match: m[0], offset: m.index });
        if (!re.global) {
          break;
        }
      }
    }
  }
  return results;
}

/**
 * Replace text in a document's body paragraphs (mutates the document).
 *
 * Performs simple text replacement within run content. Each run's text content
 * is individually searched and replaced. For cross-run matches, the paragraph
 * text is reconstructed and split back into runs.
 *
 * @param doc - The document model to modify (mutated in place).
 * @param search - String or RegExp to find.
 * @param replacement - Replacement string (supports $1, $2 etc. for RegExp).
 * @returns Number of replacements made.
 */
export function replaceText(
  doc: DocxDocument,
  search: string | RegExp,
  replacement: string
): number {
  let count = 0;
  for (const block of doc.body) {
    if (block.type !== "paragraph") {
      continue;
    }
    count += replaceInParagraph(block, search, replacement);
  }
  return count;
}

// =============================================================================
// Mail Merge
// =============================================================================

/** Regex to parse MERGEFIELD instruction: MERGEFIELD "FieldName" or MERGEFIELD FieldName */
const MERGEFIELD_RE = /^\s*MERGEFIELD\s+(?:"([^"]+)"|(\S+))/i;

/**
 * Execute a mail merge: replace all MERGEFIELD fields in the document with values from the data map.
 *
 * Fields not found in the data map are left unchanged (or optionally cleared).
 *
 * @param doc - The document to modify (mutated in place).
 * @param data - Map of field names to replacement values.
 * @param options - Optional settings.
 * @returns The number of fields replaced.
 */
export function mailMerge(
  doc: DocxDocument,
  data: Record<string, string>,
  options?: {
    /** If true, remove fields not found in data. Default: false (leave unchanged). */
    removeUnmatched?: boolean;
  }
): number {
  let count = 0;
  const removeUnmatched = options?.removeUnmatched ?? false;

  // Process body
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      count += mergeFieldsInParagraph(block, data, removeUnmatched);
    } else if (block.type === "table") {
      count += mergeFieldsInTable(block, data, removeUnmatched);
    } else if (block.type === "sdt") {
      for (const sdtChild of block.content) {
        if ("type" in sdtChild && sdtChild.type === "paragraph") {
          count += mergeFieldsInParagraph(sdtChild, data, removeUnmatched);
        } else if ("type" in sdtChild && sdtChild.type === "table") {
          count += mergeFieldsInTable(sdtChild, data, removeUnmatched);
        }
      }
    }
  }

  // Process headers and footers
  if (doc.headers) {
    for (const [, header] of doc.headers) {
      count += mergeFieldsInHeaderFooter(header.content, data, removeUnmatched);
    }
  }
  if (doc.footers) {
    for (const [, footer] of doc.footers) {
      count += mergeFieldsInHeaderFooter(footer.content, data, removeUnmatched);
    }
  }

  return count;
}

function mergeFieldsInTable(
  table: Table,
  data: Record<string, string>,
  removeUnmatched: boolean
): number {
  let count = 0;
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const block of cell.content) {
        if (block.type === "paragraph") {
          count += mergeFieldsInParagraph(block, data, removeUnmatched);
        } else if (block.type === "table") {
          count += mergeFieldsInTable(block, data, removeUnmatched);
        }
      }
    }
  }
  return count;
}

function mergeFieldsInHeaderFooter(
  content: HeaderFooterContent,
  data: Record<string, string>,
  removeUnmatched: boolean
): number {
  let count = 0;
  for (const child of content.children) {
    if (child.type === "paragraph") {
      count += mergeFieldsInParagraph(child, data, removeUnmatched);
    } else if (child.type === "table") {
      count += mergeFieldsInTable(child, data, removeUnmatched);
    }
  }
  return count;
}

function mergeFieldsInParagraph(
  para: Paragraph,
  data: Record<string, string>,
  removeUnmatched: boolean
): number {
  let count = 0;
  const children = para.children as ParagraphChild[];

  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (!("content" in child) || !Array.isArray(child.content)) {
      continue;
    }
    const run = child as Run;
    const content = run.content as RunContent[];
    for (let j = 0; j < content.length; j++) {
      const c = content[j];
      if (c.type !== "field") {
        continue;
      }
      const match = MERGEFIELD_RE.exec(c.instruction);
      if (!match) {
        continue;
      }
      const fieldName = match[1] ?? match[2];
      if (fieldName in data) {
        // Replace field with text
        content[j] = { type: "text", text: data[fieldName] } as RunContent;
        count++;
      } else if (removeUnmatched) {
        // Remove unmatched field
        content[j] = { type: "text", text: "" } as RunContent;
        count++;
      }
    }
  }
  return count;
}

/** Replace text within a single paragraph. */
function replaceInParagraph(para: Paragraph, search: string | RegExp, replacement: string): number {
  // First try simple per-run replacement
  let count = 0;
  for (const child of para.children) {
    if (!("content" in child) || !Array.isArray(child.content)) {
      continue;
    }
    for (const c of child.content) {
      if (!("type" in c) || c.type !== "text" || !("text" in c)) {
        continue;
      }
      const before = (c as { text: string }).text;
      const after =
        typeof search === "string"
          ? replaceAll(before, search, replacement)
          : before.replace(search, replacement);
      if (after !== before) {
        (c as { text: string }).text = after;
        count += typeof search === "string" ? countOccurrences(before, search) : 1;
      }
    }
  }

  // If no per-run matches, try cross-run replacement
  if (count === 0) {
    const fullText = paragraphText(para);
    const newText =
      typeof search === "string"
        ? replaceAll(fullText, search, replacement)
        : fullText.replace(search, replacement);
    if (newText !== fullText) {
      // Rebuild: put all text into first text run, clear others
      let placed = false;
      for (const child of para.children) {
        if (!("content" in child) || !Array.isArray(child.content)) {
          continue;
        }
        for (const c of child.content) {
          if (!("type" in c) || c.type !== "text" || !("text" in c)) {
            continue;
          }
          if (!placed) {
            (c as { text: string }).text = newText;
            placed = true;
          } else {
            (c as { text: string }).text = "";
          }
        }
      }
      count = typeof search === "string" ? countOccurrences(fullText, search) : 1;
    }
  }

  return count;
}

function replaceAll(str: string, search: string, replacement: string): string {
  if (search === "") {
    return str;
  }
  return str.replaceAll(search, replacement);
}

function countOccurrences(str: string, search: string): number {
  if (search === "") {
    return 0;
  }
  return str.split(search).length - 1;
}

// =============================================================================
// Patcher / Template Fill API
// =============================================================================

/** Type of content to patch into a placeholder. */
export type PatchContent =
  | { readonly type: "text"; readonly text: string }
  | { readonly type: "paragraph"; readonly children: readonly Paragraph[] }
  | { readonly type: "table"; readonly table: Table }
  | {
      readonly type: "image";
      readonly image: ImageDef;
      readonly width: number;
      readonly height: number;
    };

/** A single patch operation mapping a placeholder to replacement content. */
export interface PatchOperation {
  /** Placeholder string to find (e.g. "{{name}}"). */
  readonly placeholder: string;
  /** Content to replace the placeholder with. */
  readonly content: PatchContent;
}

/** Options for patchDocument. */
export interface PatchOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
}

/**
 * Read an existing DOCX file, replace placeholders with content, and produce a new DOCX.
 *
 * Placeholders are strings like `{{name}}` embedded in the document text.
 * They may span across multiple runs — the patcher handles cross-run matching.
 *
 * Supported patch content types:
 * - `text` — simple text replacement (preserves formatting of the first run)
 * - `paragraph` — replaces the entire paragraph containing the placeholder
 * - `table` — replaces the entire paragraph with a table
 * - `image` — replaces the placeholder with an inline image
 *
 * @param buffer - The source DOCX file as a Uint8Array.
 * @param patches - Array of patch operations to apply.
 * @param options - Optional compression settings.
 * @returns New DOCX file as a Uint8Array.
 */
export async function patchDocument(
  buffer: Uint8Array,
  patches: readonly PatchOperation[],
  options?: PatchOptions
): Promise<Uint8Array> {
  const doc = await readDocx(buffer);

  // Build lookup map for quick placeholder matching
  const patchMap = new Map<string, PatchOperation>();
  for (const patch of patches) {
    patchMap.set(patch.placeholder, patch);
  }

  // Process body content
  const newBody: BodyContent[] = [];
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      const result = patchParagraph(block, patchMap);
      if (result) {
        if (Array.isArray(result)) {
          newBody.push(...result);
        } else {
          newBody.push(result);
        }
      }
    } else if (block.type === "table") {
      patchTable(block as Table, patchMap);
      newBody.push(block);
    } else {
      newBody.push(block);
    }
  }

  // Patch headers
  if (doc.headers) {
    for (const [, headerDef] of doc.headers) {
      patchHeaderFooterContent(headerDef.content, patchMap);
    }
  }

  // Patch footers
  if (doc.footers) {
    for (const [, footerDef] of doc.footers) {
      patchHeaderFooterContent(footerDef.content, patchMap);
    }
  }

  // Add any new images from patches
  const images = doc.images ? [...doc.images] : [];
  for (const patch of patches) {
    if (patch.content.type === "image") {
      const imgContent = patch.content;
      const existing = images.find(i => i.fileName === imgContent.image.fileName);
      if (!existing) {
        images.push(imgContent.image);
      }
    }
  }

  const patched: DocxDocument = {
    ...doc,
    body: newBody,
    images: images.length > 0 ? images : undefined
  };

  return packageDocx(patched, options?.compressionLevel);
}

/** Patch a paragraph — returns replacement content or null to remove. */
function patchParagraph(
  para: Paragraph,
  patchMap: Map<string, PatchOperation>
): BodyContent | BodyContent[] | null {
  const text = paragraphText(para);

  for (const [placeholder, patch] of patchMap) {
    if (!text.includes(placeholder)) {
      continue;
    }

    switch (patch.content.type) {
      case "text": {
        replaceInParagraph(para, placeholder, patch.content.text);
        return para;
      }
      case "paragraph": {
        return patch.content.children as BodyContent[];
      }
      case "table": {
        return patch.content.table;
      }
      case "image": {
        const img = patch.content.image;
        const rId = img.rId ?? `rId_img_${img.fileName}`;
        const imgContent: InlineImageContent = {
          type: "image",
          rId,
          width: patch.content.width,
          height: patch.content.height,
          altText: img.fileName,
          name: img.fileName
        };
        const newPara: Paragraph = {
          type: "paragraph",
          properties: para.properties,
          children: [{ content: [imgContent] } as Run]
        };
        return newPara;
      }
    }
  }

  return para;
}

/** Patch text inside table cells recursively. */
function patchTable(table: Table, patchMap: Map<string, PatchOperation>): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const newContent: BodyContent[] = [];
      for (const block of cell.content) {
        if (block.type === "paragraph") {
          const result = patchParagraph(block, patchMap);
          if (result) {
            if (Array.isArray(result)) {
              newContent.push(...result);
            } else {
              newContent.push(result);
            }
          }
        } else if (block.type === "table") {
          patchTable(block as Table, patchMap);
          newContent.push(block);
        } else {
          newContent.push(block);
        }
      }
      (cell as any).content = newContent;
    }
  }
}

/** Patch text in header/footer content. */
function patchHeaderFooterContent(
  content: HeaderFooterContent,
  patchMap: Map<string, PatchOperation>
): void {
  for (const child of content.children) {
    if (child.type === "paragraph") {
      for (const [placeholder, patch] of patchMap) {
        if (patch.content.type === "text") {
          const text = paragraphText(child);
          if (text.includes(placeholder)) {
            replaceInParagraph(child, placeholder, patch.content.text);
          }
        }
      }
    }
  }
}
