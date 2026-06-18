/**
 * Run-level builder functions for DOCX documents.
 *
 * These are pure run/field/math/symbol/image/shape/chart/SDT/checkbox builders
 * that do not depend on paragraph or table builders.
 */

import type {
  Run,
  RunProperties,
  MathBlock,
  MathContent,
  CheckBox,
  FloatingImage,
  DrawingShape,
  Chart,
  ChartSeries,
  ChartAxis,
  ChartLegendPosition,
  ChartContent,
  StructuredDocumentTag,
  SdtProperties,
  Paragraph,
  Table,
  Emu,
  HexColor,
  PositionalTabAlignment,
  PositionalTabRelativeTo,
  PositionalTabLeader,
  RubyProperties,
  TextFormField,
  CheckBoxFormField,
  DropDownFormField
} from "@word/types";

// =============================================================================
// Run Builders
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

/** Create an underlined text run. */
export function underline(
  content: string,
  style?: "single" | "double" | "thick" | "dotted" | "dash" | "wave",
  properties?: Omit<RunProperties, "underline">
): Run {
  return text(content, { ...properties, underline: style ?? "single" });
}

/** Create a strikethrough text run. */
export function strikethrough(content: string, properties?: Omit<RunProperties, "strike">): Run {
  return text(content, { ...properties, strike: true });
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

/**
 * Internal helper: create a field with a fixed instruction code.
 * Used to simplify the dozens of single-instruction field builders.
 */
function simpleField(code: string, cachedValue?: string): Run {
  return field(` ${code} `, cachedValue);
}

/** Create a PAGE field (current page number). */
export function pageNumberField(cachedValue?: string): Run {
  return simpleField("PAGE", cachedValue ?? "1");
}

/** Create a NUMPAGES field (total page count). */
export function totalPagesField(cachedValue?: string): Run {
  return simpleField("NUMPAGES", cachedValue ?? "1");
}

/** Create a SECTIONPAGES field (pages in section). */
export function sectionPagesField(cachedValue?: string): Run {
  return simpleField("SECTIONPAGES", cachedValue ?? "1");
}

/** Create a SECTION field (current section number). */
export function sectionField(cachedValue?: string): Run {
  return simpleField("SECTION", cachedValue ?? "1");
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
  return simpleField("AUTHOR", cachedValue);
}

/** Create a TITLE field. */
export function titleField(cachedValue?: string): Run {
  return simpleField("TITLE", cachedValue);
}

/** Create a SUBJECT field. */
export function subjectField(cachedValue?: string): Run {
  return simpleField("SUBJECT", cachedValue);
}

/** Create a KEYWORDS field. */
export function keywordsField(cachedValue?: string): Run {
  return simpleField("KEYWORDS", cachedValue);
}

/** Create a FILENAME field. */
export function fileNameField(options?: { includePath?: boolean; cachedValue?: string }): Run {
  const instruction = options?.includePath ? " FILENAME \\p " : " FILENAME ";
  return field(instruction, options?.cachedValue);
}

/** Create a FILESIZE field. */
export function fileSizeField(cachedValue?: string): Run {
  return simpleField("FILESIZE", cachedValue);
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
  /**
   * Right-align page numbers.
   *
   * NOTE: Right-aligned page numbers are the TOC default — they come from the
   * right-aligned tab stop in the TOC paragraph styles, not from a field
   * switch. The `\z` switch does NOT mean "right align": per ECMA-376 it
   * *hides* the tab leader and page numbers in Web layout view. Mapping this
   * option to `\z` therefore broke the layout, so we no longer emit it.
   */
  rightAlignedPageNumbers?: boolean;
  /**
   * Tab leader style between an entry and its page number.
   *
   * NOTE: The dotted leader is already the TOC default (a tab stop with dot
   * leader defined by the TOC paragraph styles). The TOC field has no switch
   * for choosing the leader glyph — the `\p` switch sets the *separator
   * character* (replacing the tab entirely), which would DISABLE the leader
   * dots and the right-aligned page number. We therefore do not translate
   * this option into `\p`; the leader is controlled by the TOC styles.
   */
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
  // Intentionally NOT emitting `\z` for rightAlignedPageNumbers — see note above.
  // Intentionally NOT emitting `\p` for tabLeader — see the field note above.
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

/** Create a FORMTEXT field (legacy text form field) with full properties. */
export function formTextField(options?: {
  name?: string;
  default?: string;
  maxLength?: number;
  format?: string;
  helpText?: string;
  statusText?: string;
  enabled?: boolean;
  cachedValue?: string;
}): Run {
  const formField: TextFormField = {
    type: "text",
    ...(options?.name !== undefined && { name: options.name }),
    ...(options?.default !== undefined && { default: options.default }),
    ...(options?.maxLength !== undefined && { maxLength: options.maxLength }),
    ...(options?.format !== undefined && { format: options.format }),
    ...(options?.helpText !== undefined && { helpText: options.helpText }),
    ...(options?.statusText !== undefined && { statusText: options.statusText }),
    ...(options?.enabled !== undefined && { enabled: options.enabled })
  };
  return {
    content: [
      {
        type: "field",
        instruction: " FORMTEXT ",
        cachedValue: options?.cachedValue,
        formField
      }
    ]
  };
}

/** Create a FORMCHECKBOX field (legacy checkbox form field) with full properties. */
export function formCheckboxField(options?: {
  name?: string;
  checked?: boolean;
  default?: boolean;
  size?: number;
  cachedValue?: string;
}): Run {
  const formField: CheckBoxFormField = {
    type: "checkBox",
    ...(options?.name !== undefined && { name: options.name }),
    ...(options?.checked !== undefined && { checked: options.checked }),
    ...(options?.default !== undefined && { default: options.default }),
    ...(options?.size !== undefined && { size: options.size })
  };
  return {
    content: [
      {
        type: "field",
        instruction: " FORMCHECKBOX ",
        cachedValue: options?.cachedValue,
        formField
      }
    ]
  };
}

/** Create a FORMDROPDOWN field (legacy dropdown form field) with full properties. */
export function formDropdownField(options?: {
  name?: string;
  entries?: readonly string[];
  default?: number;
  helpText?: string;
  statusText?: string;
  enabled?: boolean;
  cachedValue?: string;
}): Run {
  const formField: DropDownFormField = {
    type: "dropDown",
    ...(options?.name !== undefined && { name: options.name }),
    ...(options?.entries !== undefined && { entries: options.entries }),
    ...(options?.default !== undefined && { default: options.default }),
    ...(options?.helpText !== undefined && { helpText: options.helpText }),
    ...(options?.statusText !== undefined && { statusText: options.statusText }),
    ...(options?.enabled !== undefined && { enabled: options.enabled })
  };
  return {
    content: [
      {
        type: "field",
        instruction: " FORMDROPDOWN ",
        cachedValue: options?.cachedValue,
        formField
      }
    ]
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

/**
 * Create a math phantom (an expression that takes up space).
 *
 * Note: in OOXML the phantom base is *shown* by default. To make the classic
 * "occupies space but invisible" phantom pass `{ show: false }`; passing only
 * `transparent: true` is not sufficient to hide the base in Word.
 */
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

// =============================================================================
// Symbol / Image / Shape / Chart / SDT / CheckBox Builders
// =============================================================================

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
