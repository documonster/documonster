/**
 * Run-level builder functions for DOCX documents.
 *
 * These are pure run/field/math/symbol/image/shape/chart/SDT/checkbox builders
 * that do not depend on paragraph or table builders.
 */

import { isRecordArgument } from "@utils/object";
import {
  checkArray,
  checkBoolean,
  checkBooleans,
  checkBoundedString,
  CHART_TYPES,
  checkChar,
  checkChartLegend,
  checkChartSeries,
  checkFieldName,
  checkHeadingLevel,
  checkHeadingRange,
  checkInteger,
  checkMathFractionType,
  checkMathLimitType,
  checkMathPosition,
  checkMathRows,
  checkMathVerticalAlign,
  checkNodeArray,
  checkOneOf,
  checkPositionalTab,
  checkRecord,
  checkRubyProperties,
  checkSdtProperties,
  checkRgbColor,
  checkRunProperties,
  checkString,
  checkUnderlineStyle,
  under
} from "@word/core/argument-checks";
import { assertDrawingNumber, assertDrawingNumbers } from "@word/core/drawing-geometry";
import type {
  Run,
  RunProperties,
  UnderlineStyle,
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

/**
 * A text run, its arguments checked in the name of `where` — the public
 * builder that was called, so a wrong value is reported against it.
 */
export function checkedTextRun(
  where: string,
  content: string,
  properties: RunProperties | undefined
): Run {
  checkString(content, "content", where);
  checkRunProperties(properties, "properties", where);
  return { properties, content: [{ type: "text", text: content }] };
}

/** Create a text run. */
export function text(content: string, properties?: RunProperties): Run {
  return checkedTextRun("Build.text", content, properties);
}

/** Create a bold text run. */
export function bold(content: string, properties?: Omit<RunProperties, "bold">): Run {
  return checkedTextRun("Build.bold", content, { ...properties, bold: true });
}

/** Create an italic text run. */
export function italic(content: string, properties?: Omit<RunProperties, "italic">): Run {
  return checkedTextRun("Build.italic", content, { ...properties, italic: true });
}

/** Create an underlined text run. */
export function underline(content: string, properties?: RunProperties): Run;
/**
 * @deprecated Pass the style in the properties: `underline(content, { ...properties, underline: style })`.
 * Removed in 1.0.
 */
export function underline(
  content: string,
  style: "single" | "double" | "thick" | "dotted" | "dash" | "wave" | undefined,
  properties?: Omit<RunProperties, "underline">
): Run;
export function underline(
  content: string,
  styleOrProperties?: string | RunProperties,
  properties?: Omit<RunProperties, "underline">
): Run {
  // Anything but an options object is the deprecated style, checked as the style.
  if (
    arguments.length > 2 ||
    (styleOrProperties !== undefined && !isRecordArgument(styleOrProperties))
  ) {
    checkUnderlineStyle(styleOrProperties, "style", "Build.underline");
    return checkedTextRun("Build.underline", content, {
      ...properties,
      underline: (styleOrProperties as UnderlineStyle | undefined) ?? "single"
    });
  }
  // An underline the properties do not choose is a single one.
  return checkedTextRun("Build.underline", content, {
    ...styleOrProperties,
    underline: styleOrProperties?.underline ?? "single"
  });
}

/** Create a strikethrough text run. */
export function strikethrough(content: string, properties?: Omit<RunProperties, "strike">): Run {
  return checkedTextRun("Build.strikethrough", content, { ...properties, strike: true });
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
  checkPositionalTab(options, "options", "Build.positionalTab");
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
  const where = "Build.ruby";
  checkRubyProperties(properties, "properties", where);
  const normalize = (v: string | Run | readonly Run[], field: string): readonly Run[] => {
    if (typeof v === "string") {
      return [checkedTextRun(where, v, undefined)];
    }
    if (Array.isArray(v)) {
      checkNodeArray(v, field, where);
      return v as readonly Run[];
    }
    checkRecord(v, field, where);
    return [v as Run];
  };
  const base = normalize(baseText, "baseText");
  const annotation = normalize(rubyText, "rubyText");
  return {
    content: [
      {
        type: "ruby",
        properties,
        baseText: base,
        rubyText: annotation
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
export function field(instruction: string, options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function field(instruction: string, cachedValue: string): Run;
export function field(instruction: string, options?: FieldOptions | string): Run {
  checkString(instruction, "instruction", "Build.field");
  return checkedField("Build.field", instruction, cachedValueOf(options, "Build.field"));
}

/** A field run, its cached value checked in the name of `where`. */
function checkedField(where: string, instruction: string, cachedValue: unknown): Run {
  checkString(cachedValue, "cachedValue", where, { optional: true });
  return { content: [{ type: "field", instruction, cachedValue: cachedValue as string }] };
}

/**
 * Internal helper: create a field with a fixed instruction code.
 * Used to simplify the dozens of single-instruction field builders.
 */
function simpleField(where: string, code: string, cachedValue?: string): Run {
  return checkedField(where, ` ${code} `, cachedValue);
}

/** `options`, when given, is an object; its `cachedValue` is checked by the field. */
function checkFieldOptions(options: unknown, where: string): void {
  checkRecord(options, "options", where, { optional: true });
}

/** What every field takes: the result shown until the field is updated. */
export interface FieldOptions {
  /** The result shown until the field is updated. */
  readonly cachedValue?: string;
}

/** The cached value of a field call: from `{ cachedValue }`, or the deprecated positional string. */
function cachedValueOf(
  options: FieldOptions | string | undefined,
  where: string
): string | undefined {
  if (isRecordArgument(options)) {
    checkFieldOptions(options, where);
    return options.cachedValue;
  }
  // Anything else is the positional form, which the field checks as `cachedValue`, as it always has.
  return options;
}

/** Create a PAGE field (current page number). */
export function pageNumberField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function pageNumberField(cachedValue: string): Run;
export function pageNumberField(options?: FieldOptions | string): Run {
  const where = "Build.pageNumberField";
  return simpleField(where, "PAGE", cachedValueOf(options, where) ?? "1");
}

/** Create a NUMPAGES field (total page count). */
export function totalPagesField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function totalPagesField(cachedValue: string): Run;
export function totalPagesField(options?: FieldOptions | string): Run {
  const where = "Build.totalPagesField";
  return simpleField(where, "NUMPAGES", cachedValueOf(options, where) ?? "1");
}

/** Create a SECTIONPAGES field (pages in section). */
export function sectionPagesField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function sectionPagesField(cachedValue: string): Run;
export function sectionPagesField(options?: FieldOptions | string): Run {
  const where = "Build.sectionPagesField";
  return simpleField(where, "SECTIONPAGES", cachedValueOf(options, where) ?? "1");
}

/** Create a SECTION field (current section number). */
export function sectionField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function sectionField(cachedValue: string): Run;
export function sectionField(options?: FieldOptions | string): Run {
  const where = "Build.sectionField";
  return simpleField(where, "SECTION", cachedValueOf(options, where) ?? "1");
}

/** Options for {@link dateField} and {@link timeField}. */
export interface DateTimeFieldOptions {
  /** Date/time picture (the `\@` switch). Default: `"yyyy-MM-dd"` for a date, `"HH:mm:ss"` for a time. */
  readonly format?: string;
  /** The result shown until the field is updated. */
  readonly cachedValue?: string;
}

/** Create a DATE field: `dateField({ format: "d MMMM yyyy" })`. */
export function dateField(options?: DateTimeFieldOptions): Run;
/**
 * @deprecated Pass an options object: `dateField({ format, cachedValue })`. Removed in 1.0.
 */
export function dateField(format: string | undefined, cachedValue?: string): Run;
export function dateField(
  formatOrOptions?: string | DateTimeFieldOptions,
  cachedValue?: string
): Run {
  const where = "Build.dateField";
  const named = isRecordArgument(formatOrOptions);
  const { format, cachedValue: cached } = named
    ? formatOrOptions
    : { format: formatOrOptions, cachedValue };
  checkString(format, named ? "options.format" : "format", where, { optional: true });
  const fmt = format ?? "yyyy-MM-dd";
  return checkedField(where, ` DATE \\@ "${fmt}" `, cached);
}

/** Create a SEQ (sequence) field for numbering figures, tables, etc. */
export function sequenceField(
  identifier: string,
  options?: { cachedValue?: string; hide?: boolean }
): Run {
  const where = "Build.sequenceField";
  checkFieldName(identifier, "identifier", where);
  checkFieldOptions(options, where);
  checkBoolean(options?.hide, "options.hide", where, { optional: true });
  let instruction = ` SEQ ${identifier} `;
  if (options?.hide) {
    instruction += "\\h ";
  }
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create a TIME field (current time): `timeField({ format: "HH:mm" })`. */
export function timeField(options?: DateTimeFieldOptions): Run;
/**
 * @deprecated Pass an options object: `timeField({ format, cachedValue })`. Removed in 1.0.
 */
export function timeField(format: string | undefined, cachedValue?: string): Run;
export function timeField(
  formatOrOptions?: string | DateTimeFieldOptions,
  cachedValue?: string
): Run {
  const where = "Build.timeField";
  const named = isRecordArgument(formatOrOptions);
  const { format, cachedValue: cached } = named
    ? formatOrOptions
    : { format: formatOrOptions, cachedValue };
  checkString(format, named ? "options.format" : "format", where, { optional: true });
  const fmt = format ?? "HH:mm:ss";
  return checkedField(where, ` TIME \\@ "${fmt}" `, cached);
}

/** Create an AUTHOR field. */
export function authorField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function authorField(cachedValue: string): Run;
export function authorField(options?: FieldOptions | string): Run {
  const where = "Build.authorField";
  return simpleField(where, "AUTHOR", cachedValueOf(options, where));
}

/** Create a TITLE field. */
export function titleField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function titleField(cachedValue: string): Run;
export function titleField(options?: FieldOptions | string): Run {
  const where = "Build.titleField";
  return simpleField(where, "TITLE", cachedValueOf(options, where));
}

/** Create a SUBJECT field. */
export function subjectField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function subjectField(cachedValue: string): Run;
export function subjectField(options?: FieldOptions | string): Run {
  const where = "Build.subjectField";
  return simpleField(where, "SUBJECT", cachedValueOf(options, where));
}

/** Create a KEYWORDS field. */
export function keywordsField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function keywordsField(cachedValue: string): Run;
export function keywordsField(options?: FieldOptions | string): Run {
  const where = "Build.keywordsField";
  return simpleField(where, "KEYWORDS", cachedValueOf(options, where));
}

/** Create a FILENAME field. */
export function fileNameField(options?: { includePath?: boolean; cachedValue?: string }): Run {
  const where = "Build.fileNameField";
  checkBooleans(options, ["includePath"], "options", where);
  const instruction = options?.includePath ? " FILENAME \\p " : " FILENAME ";
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create a FILESIZE field. */
export function fileSizeField(options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function fileSizeField(cachedValue: string): Run;
export function fileSizeField(options?: FieldOptions | string): Run {
  const where = "Build.fileSizeField";
  return simpleField(where, "FILESIZE", cachedValueOf(options, where));
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
  const where = "Build.styleRefField";
  checkString(styleName, "styleName", where);
  checkBooleans(
    options,
    ["fromBottom", "insertParagraphNumber", "insertPosition", "suppressNonNumeric"],
    "options",
    where
  );
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
  return checkedField(where, instruction, options?.cachedValue);
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
  const where = "Build.refField";
  checkFieldName(bookmarkName, "bookmarkName", where);
  checkBooleans(
    options,
    [
      "insertNumber",
      "hyperlink",
      "paragraphNumber",
      "relativeParagraphNumber",
      "fullContext",
      "suppressNonDelimiter"
    ],
    "options",
    where
  );
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
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create a PAGEREF field (references page of bookmark). */
export function pageRefField(
  bookmarkName: string,
  options?: { hyperlink?: boolean; relativePosition?: boolean; cachedValue?: string }
): Run {
  const where = "Build.pageRefField";
  checkFieldName(bookmarkName, "bookmarkName", where);
  checkBooleans(options, ["hyperlink", "relativePosition"], "options", where);
  let instruction = ` PAGEREF ${bookmarkName} `;
  if (options?.hyperlink) {
    instruction += "\\h ";
  }
  if (options?.relativePosition) {
    instruction += "\\p ";
  }
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create a NOTEREF field (references a footnote/endnote). */
export function noteRefField(
  bookmarkName: string,
  options?: { hyperlink?: boolean; insertNumberFormat?: boolean; cachedValue?: string }
): Run {
  const where = "Build.noteRefField";
  checkFieldName(bookmarkName, "bookmarkName", where);
  checkBooleans(options, ["hyperlink", "insertNumberFormat"], "options", where);
  let instruction = ` NOTEREF ${bookmarkName} `;
  if (options?.hyperlink) {
    instruction += "\\h ";
  }
  if (options?.insertNumberFormat) {
    instruction += "\\f ";
  }
  return checkedField(where, instruction, options?.cachedValue);
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
  const where = "Build.hyperlinkField";
  checkString(target, "target", where);
  checkBooleans(options, ["newWindow"], "options", where);
  for (const key of ["anchor", "displayText", "tooltip"] as const) {
    checkString(options?.[key], `options.${key}`, where, { optional: true });
  }
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
  return checkedField(where, instruction, options?.cachedValue ?? options?.displayText);
}

/** Create a QUOTE field (literal text). */
export function quoteField(text: string, options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function quoteField(text: string, cachedValue: string): Run;
export function quoteField(text: string, options?: FieldOptions | string): Run {
  const where = "Build.quoteField";
  checkString(text, "text", where);
  return checkedField(where, ` QUOTE "${text}" `, cachedValueOf(options, where) ?? text);
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
  const where = "Build.tocField";
  checkBooleans(
    options,
    ["hyperlink", "rightAlignedPageNumbers", "noPageNumbers"],
    "options",
    where
  );
  checkHeadingRange(options?.headingLevels, "options.headingLevels", where, { optional: true });
  if (options?.styles !== undefined) {
    checkArray(options.styles, "options.styles", where);
    options.styles.forEach((style, i) => checkString(style, `options.styles[${i}]`, where));
  }
  checkString(options?.tcLevels, "options.tcLevels", where, { optional: true });
  checkOneOf(options?.tabLeader, TOC_TAB_LEADERS, "options.tabLeader", where, { optional: true });
  checkString(options?.captionIdentifier, "options.captionIdentifier", where, { optional: true });
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
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create a TC (table of contents entry) field. */
export function tcField(
  text: string,
  options?: { level?: number; suppressPageNumber?: boolean; cachedValue?: string }
): Run {
  const where = "Build.tcField";
  checkString(text, "text", where);
  checkBooleans(options, ["suppressPageNumber"], "options", where);
  if (options?.level !== undefined) {
    checkHeadingLevel(options.level, "options.level", where);
  }
  let instruction = ` TC "${text}" `;
  if (options?.level !== undefined) {
    instruction += `\\l ${options.level} `;
  }
  if (options?.suppressPageNumber) {
    instruction += "\\n ";
  }
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create an XE (index entry) field. */
export function indexEntryField(
  text: string,
  options?: { bold?: boolean; italic?: boolean; cachedValue?: string }
): Run {
  const where = "Build.indexEntryField";
  checkString(text, "text", where);
  checkBooleans(options, ["bold", "italic"], "options", where);
  let instruction = ` XE "${text}" `;
  if (options?.bold) {
    instruction += "\\b ";
  }
  if (options?.italic) {
    instruction += "\\i ";
  }
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create an INDEX field (renders index from XE entries). */
export function indexField(options?: {
  bookmark?: string;
  columns?: number;
  entryType?: string;
  cachedValue?: string;
}): Run {
  const where = "Build.indexField";
  checkFieldOptions(options, where);
  if (options?.bookmark !== undefined) {
    checkFieldName(options.bookmark, "options.bookmark", where);
  }
  checkInteger(options?.columns, "options.columns", where, { optional: true, min: 1 });
  checkString(options?.entryType, "options.entryType", where, { optional: true });
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
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create an IF field (conditional content). */
export function ifField(
  condition: string,
  trueText: string,
  falseText: string,
  options?: FieldOptions
): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function ifField(
  condition: string,
  trueText: string,
  falseText: string,
  cachedValue: string
): Run;
export function ifField(
  condition: string,
  trueText: string,
  falseText: string,
  options?: FieldOptions | string
): Run {
  const where = "Build.ifField";
  const cachedValue = cachedValueOf(options, where);
  checkString(condition, "condition", where);
  checkString(trueText, "trueText", where);
  checkString(falseText, "falseText", where);
  return checkedField(where, ` IF ${condition} "${trueText}" "${falseText}" `, cachedValue);
}

/** Create an INCLUDETEXT field (includes external file content). */
export function includeTextField(
  filePath: string,
  options?: { bookmark?: string; cachedValue?: string }
): Run {
  const where = "Build.includeTextField";
  checkString(filePath, "filePath", where);
  checkFieldOptions(options, where);
  if (options?.bookmark !== undefined) {
    checkFieldName(options.bookmark, "options.bookmark", where);
  }
  let instruction = ` INCLUDETEXT "${filePath}" `;
  if (options?.bookmark) {
    instruction += `${options.bookmark} `;
  }
  return checkedField(where, instruction, options?.cachedValue);
}

/** Create an INCLUDEPICTURE field. */
export function includePictureField(filePath: string, options?: FieldOptions): Run;
/** @deprecated Pass `{ cachedValue }`. Removed in 1.0. */
export function includePictureField(filePath: string, cachedValue: string): Run;
export function includePictureField(filePath: string, options?: FieldOptions | string): Run {
  const where = "Build.includePictureField";
  checkString(filePath, "filePath", where);
  return checkedField(where, ` INCLUDEPICTURE "${filePath}" `, cachedValueOf(options, where));
}

/**
 * Create a FORMTEXT field (legacy text form field) with full properties.
 *
 * The field shows `cachedValue` — what has been typed into it — and, without
 * one, its `default`, the way Word fills a new field with its default text.
 */
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
  const where = "Build.formTextField";
  checkFormFieldCommon(options, where);
  checkString(options?.default, "options.default", where, { optional: true });
  checkInteger(options?.maxLength, "options.maxLength", where, { optional: true, min: 0 });
  checkString(options?.format, "options.format", where, { optional: true });
  checkString(options?.cachedValue, "options.cachedValue", where, { optional: true });
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
        cachedValue: options?.cachedValue ?? options?.default,
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
  const where = "Build.formCheckboxField";
  checkFormFieldCommon(options, where);
  checkBooleans(options, ["checked", "default"], "options", where);
  // The box size is `w:size`, half-points.
  checkInteger(options?.size, "options.size", where, { optional: true, min: 0 });
  checkString(options?.cachedValue, "options.cachedValue", where, { optional: true });
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
  const where = "Build.formDropdownField";
  checkFormFieldCommon(options, where);
  if (options?.entries !== undefined) {
    checkArray(options.entries, "options.entries", where);
    options.entries.forEach((e, i) => checkString(e, `options.entries[${i}]`, where));
  }
  // The index of the selected entry.
  checkInteger(options?.default, "options.default", where, {
    optional: true,
    min: 0,
    ...(options?.entries !== undefined && { max: options.entries.length - 1 })
  });
  checkString(options?.cachedValue, "options.cachedValue", where, { optional: true });
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

/**
 * What every legacy form field shares: `w:name` (`ST_FFName`, at most 65
 * characters), the help and status texts (`ST_FFHelpTextVal`, 256;
 * `ST_FFStatusTextVal`, 140) and `enabled`.
 */
function checkFormFieldCommon(
  options:
    | { name?: unknown; helpText?: unknown; statusText?: unknown; enabled?: unknown }
    | undefined,
  where: string
): void {
  checkRecord(options, "options", where, { optional: true });
  checkBoundedString(options?.name, 65, "options.name", where, { optional: true });
  checkBoundedString(options?.helpText, 256, "options.helpText", where, { optional: true });
  checkBoundedString(options?.statusText, 140, "options.statusText", where, { optional: true });
  checkBoolean(options?.enabled, "options.enabled", where, { optional: true });
}

const TOC_TAB_LEADERS: ReadonlySet<string> = new Set([".", "-", "_", " "]);

// =============================================================================
// Math Builders
// =============================================================================

/** Each named argument is math content — an array of `Build.math*` nodes. */
function checkMath(
  where: string,
  args: Record<string, unknown>,
  optionalKeys: string[] = []
): void {
  for (const [field, value] of Object.entries(args)) {
    checkNodeArray(value, field, where, { optional: optionalKeys.includes(field) });
  }
}

/** Create a math block. */
export function mathBlock(content: readonly MathContent[]): MathBlock {
  checkMath("Build.mathBlock", { content });
  return { type: "math", content };
}

/** Create a math text run. */
export function mathRun(
  mathText: string,
  properties?: { italic?: boolean; bold?: boolean; font?: string }
): MathContent {
  const where = "Build.mathRun";
  checkString(mathText, "mathText", where);
  checkBooleans(properties, ["italic", "bold"], "properties", where);
  checkString(properties?.font, "properties.font", where, { optional: true });
  return { type: "mathRun", text: mathText, properties };
}

/** Options for `Build.mathFraction`. */
export interface MathFractionOptions {
  /** How the fraction is drawn (`m:type`). Default: a horizontal bar (`"bar"`). */
  readonly fractionType?: "bar" | "skw" | "lin" | "noBar";
}

/** Create a math fraction: `mathFraction(numerator, denominator, { fractionType })`. */
export function mathFraction(
  numerator: readonly MathContent[],
  denominator: readonly MathContent[],
  options?: MathFractionOptions
): MathContent;
/**
 * @deprecated Pass the type in options: `mathFraction(numerator, denominator, { fractionType })`.
 * Removed in 1.0.
 */
export function mathFraction(
  numerator: readonly MathContent[],
  denominator: readonly MathContent[],
  fractionType: "bar" | "skw" | "lin" | "noBar" | undefined
): MathContent;
export function mathFraction(
  numerator: readonly MathContent[],
  denominator: readonly MathContent[],
  typeOrOptions?: "bar" | "skw" | "lin" | "noBar" | MathFractionOptions
): MathContent {
  const where = "Build.mathFraction";
  checkMath(where, { numerator, denominator });
  const fractionType = optionOrPositional(typeOrOptions, "fractionType");
  checkMathFractionType(fractionType.value, fractionType.field, where);
  return {
    type: "mathFraction",
    fractionType: fractionType.value as MathFractionOptions["fractionType"],
    numerator,
    denominator
  };
}

/** Create a math square root. */
export function mathSqrt(content: readonly MathContent[]): MathContent {
  checkMath("Build.mathSqrt", { content });
  return { type: "mathRadical", content, hideDegree: true };
}

/** Create a math nth root. */
export function mathRoot(
  degree: readonly MathContent[],
  content: readonly MathContent[]
): MathContent {
  checkMath("Build.mathRoot", { degree, content });
  return { type: "mathRadical", degree, content };
}

/** The limits of an n-ary operator (`Build.mathSum`, `mathIntegral`, `mathProduct`, `mathNary`). */
export interface MathNaryOptions {
  /** The lower limit (below the operator, or its subscript). */
  readonly sub?: readonly MathContent[];
  /** The upper limit (above the operator, or its superscript). */
  readonly sup?: readonly MathContent[];
}

/**
 * The limits of an n-ary operator from its options, or from the deprecated
 * positional `sub` / `sup`: a second argument that is neither an array nor
 * `undefined` is the options object.
 */
function naryLimits(
  where: string,
  content: readonly MathContent[],
  subOrOptions: readonly MathContent[] | MathNaryOptions | undefined,
  sup: readonly MathContent[] | undefined
): MathNaryOptions {
  if (subOrOptions === undefined || Array.isArray(subOrOptions)) {
    checkMath(where, { content, sub: subOrOptions, sup }, ["sub", "sup"]);
    return { sub: subOrOptions as readonly MathContent[] | undefined, sup };
  }
  checkMath(where, { content });
  checkRecord(subOrOptions, "options", where);
  const options = subOrOptions as MathNaryOptions;
  checkNodeArray(options.sub, "options.sub", where, { optional: true });
  checkNodeArray(options.sup, "options.sup", where, { optional: true });
  return { sub: options.sub, sup: options.sup };
}

/** Create a math summation: `mathSum(content, { sub, sup })`. */
export function mathSum(content: readonly MathContent[], options?: MathNaryOptions): MathContent;
/** @deprecated Pass the limits in options: `mathSum(content, { sub, sup })`. Removed in 1.0. */
export function mathSum(
  content: readonly MathContent[],
  sub: readonly MathContent[] | undefined,
  sup?: readonly MathContent[]
): MathContent;
export function mathSum(
  content: readonly MathContent[],
  subOrOptions?: readonly MathContent[] | MathNaryOptions,
  sup?: readonly MathContent[]
): MathContent {
  const limits = naryLimits("Build.mathSum", content, subOrOptions, sup);
  return { type: "mathNary", char: "\u2211", sub: limits.sub, sup: limits.sup, content };
}

/** Create a math integral: `mathIntegral(content, { sub, sup })`. */
export function mathIntegral(
  content: readonly MathContent[],
  options?: MathNaryOptions
): MathContent;
/** @deprecated Pass the limits in options: `mathIntegral(content, { sub, sup })`. Removed in 1.0. */
export function mathIntegral(
  content: readonly MathContent[],
  sub: readonly MathContent[] | undefined,
  sup?: readonly MathContent[]
): MathContent;
export function mathIntegral(
  content: readonly MathContent[],
  subOrOptions?: readonly MathContent[] | MathNaryOptions,
  sup?: readonly MathContent[]
): MathContent {
  const limits = naryLimits("Build.mathIntegral", content, subOrOptions, sup);
  return { type: "mathNary", char: "\u222B", sub: limits.sub, sup: limits.sup, content };
}

/** Create a math product: `mathProduct(content, { sub, sup })`. */
export function mathProduct(
  content: readonly MathContent[],
  options?: MathNaryOptions
): MathContent;
/** @deprecated Pass the limits in options: `mathProduct(content, { sub, sup })`. Removed in 1.0. */
export function mathProduct(
  content: readonly MathContent[],
  sub: readonly MathContent[] | undefined,
  sup?: readonly MathContent[]
): MathContent;
export function mathProduct(
  content: readonly MathContent[],
  subOrOptions?: readonly MathContent[] | MathNaryOptions,
  sup?: readonly MathContent[]
): MathContent {
  const limits = naryLimits("Build.mathProduct", content, subOrOptions, sup);
  return { type: "mathNary", char: "\u220F", sub: limits.sub, sup: limits.sup, content };
}

/** Create a math superscript. */
export function mathSuperScript(
  base: readonly MathContent[],
  superScript: readonly MathContent[]
): MathContent {
  checkMath("Build.mathSuperscript", { base, superScript });
  return { type: "mathSuperScript", base, superScript };
}

/** Create a math subscript. */
export function mathSubScript(
  base: readonly MathContent[],
  subScript: readonly MathContent[]
): MathContent {
  checkMath("Build.mathSubscript", { base, subScript });
  return { type: "mathSubScript", base, subScript };
}

/** Create a math sub-superscript. */
export function mathSubSuperScript(
  base: readonly MathContent[],
  subScript: readonly MathContent[],
  superScript: readonly MathContent[]
): MathContent {
  checkMath("Build.mathSubSuperscript", { base, subScript, superScript });
  return { type: "mathSubSuperScript", base, subScript, superScript };
}

/** Create a math pre-sub-superscript (subscript/superscript before the base). */
export function mathPreSubSuperScript(
  base: readonly MathContent[],
  preSubScript: readonly MathContent[],
  preSuperScript: readonly MathContent[]
): MathContent {
  checkMath("Build.mathPreSubSuperscript", { base, preSubScript, preSuperScript });
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
  content: readonly MathContent[],
  options?: {
    show?: boolean;
    zeroWidth?: boolean;
    zeroAscent?: boolean;
    zeroDescent?: boolean;
    transparent?: boolean;
  }
): MathContent {
  const where = "Build.mathPhantom";
  checkMath(where, { content });
  checkBooleans(
    options,
    ["show", "zeroWidth", "zeroAscent", "zeroDescent", "transparent"],
    "options",
    where
  );
  return { type: "mathPhantom", content, ...options };
}

/** Create a math group character (e.g. a horizontal brace over an expression). */
export function mathGroupChar(
  base: readonly MathContent[],
  options?: {
    char?: string;
    position?: "top" | "bottom";
    verticalAlign?: "top" | "center" | "bottom";
  }
): MathContent {
  const where = "Build.mathGroupChar";
  checkMath(where, { base });
  checkRecord(options, "options", where, { optional: true });
  checkChar(options?.char, "options.char", where, { optional: true });
  checkMathPosition(options?.position, "options.position", where);
  checkMathVerticalAlign(options?.verticalAlign, "options.verticalAlign", where);
  return { type: "mathGroupChar", base, ...options };
}

/** Create a math border box (draw borders around / strike through an expression). */
export function mathBorderBox(
  content: readonly MathContent[],
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
  const where = "Build.mathBorderBox";
  checkMath(where, { content });
  checkBooleans(
    options,
    [
      "hideTop",
      "hideBottom",
      "hideLeft",
      "hideRight",
      "strikeBlTr",
      "strikeTlBr",
      "strikeH",
      "strikeV"
    ],
    "options",
    where
  );
  return { type: "mathBorderBox", content, ...options };
}

/** Create a math delimiter (parentheses, brackets, etc.). */
export function mathDelimiter(
  content: readonly (readonly MathContent[])[],
  options?: { beginChar?: string; endChar?: string; separatorChar?: string }
): MathContent {
  const where = "Build.mathDelimiter";
  checkMathRows(content, 1, "content", where);
  checkRecord(options, "options", where, { optional: true });
  for (const key of ["beginChar", "endChar", "separatorChar"] as const) {
    // An empty delimiter character means "none".
    checkChar(options?.[key], `options.${key}`, where, { optional: true, empty: true });
  }
  return {
    type: "mathDelimiter",
    beginChar: options?.beginChar,
    endChar: options?.endChar,
    separatorChar: options?.separatorChar,
    content
  };
}

/** Create an n-ary operator with any character: `mathNary("∐", content, { sub, sup })`. */
export function mathNary(
  char: string,
  content: readonly MathContent[],
  options?: MathNaryOptions
): MathContent;
/** @deprecated Pass the limits in options: `mathNary(char, content, { sub, sup })`. Removed in 1.0. */
export function mathNary(
  char: string,
  content: readonly MathContent[],
  sub: readonly MathContent[] | undefined,
  sup?: readonly MathContent[]
): MathContent;
export function mathNary(
  char: string,
  content: readonly MathContent[],
  subOrOptions?: readonly MathContent[] | MathNaryOptions,
  sup?: readonly MathContent[]
): MathContent {
  checkChar(char, "char", "Build.mathNary");
  const limits = naryLimits("Build.mathNary", content, subOrOptions, sup);
  return { type: "mathNary", char, sub: limits.sub, sup: limits.sup, content };
}

/** Create a math function (sin, cos, lim, etc.). */
export function mathFunction(
  name: readonly MathContent[],
  content: readonly MathContent[]
): MathContent {
  checkMath("Build.mathFunction", { name, content });
  return { type: "mathFunction", name, content };
}

/** Options for `Build.mathLimit`. */
export interface MathLimitOptions {
  /** Whether the limit sits above (`"upper"`) or below (`"lower"`) the base. Default: `"lower"`. */
  readonly limitType?: "upper" | "lower";
}

/** Create a math limit: `mathLimit(base, limit, { limitType })`. */
export function mathLimit(
  base: readonly MathContent[],
  limit: readonly MathContent[],
  options?: MathLimitOptions
): MathContent;
/** @deprecated Pass the type in options: `mathLimit(base, limit, { limitType })`. Removed in 1.0. */
export function mathLimit(
  base: readonly MathContent[],
  limit: readonly MathContent[],
  limitType: "upper" | "lower" | undefined
): MathContent;
export function mathLimit(
  base: readonly MathContent[],
  limit: readonly MathContent[],
  typeOrOptions?: "upper" | "lower" | MathLimitOptions
): MathContent {
  const where = "Build.mathLimit";
  checkMath(where, { base, limit });
  const given = optionOrPositional(typeOrOptions, "limitType");
  // The default stands in for an absent value only: `null` is a wrong one, refused as before.
  const limitType = given.value === undefined ? "lower" : given.value;
  checkMathLimitType(limitType, given.field, where);
  return { type: "mathLimit", base, limit, limitType: limitType as "upper" | "lower" };
}

/** Create a math matrix. */
export function mathMatrix(rows: readonly (readonly (readonly MathContent[])[])[]): MathContent {
  checkMathRows(rows, 2, "rows", "Build.mathMatrix");
  return { type: "mathMatrix", rows };
}

/** Options for `Build.mathAccent`. */
export interface MathAccentOptions {
  /** The accent character (`m:chr`). Default: a circumflex (`"\u0302"`), as Word draws one. */
  readonly char?: string;
}

/** Create a math accent (hat, tilde, …): `mathAccent(content, { char })`. */
export function mathAccent(
  content: readonly MathContent[],
  options?: MathAccentOptions
): MathContent;
/** @deprecated Pass the character in options: `mathAccent(content, { char })`. Removed in 1.0. */
export function mathAccent(content: readonly MathContent[], char: string | undefined): MathContent;
export function mathAccent(
  content: readonly MathContent[],
  charOrOptions?: string | MathAccentOptions
): MathContent {
  const where = "Build.mathAccent";
  checkMath(where, { content });
  const char = optionOrPositional(charOrOptions, "char");
  checkChar(char.value, char.field, where, { optional: true });
  return { type: "mathAccent", content, char: char.value as string | undefined };
}

/**
 * The one setting of a builder that took it positionally before it took an options object: from
 * `options[key]`, or the deprecated positional value as it is. Anything but an options object is
 * the positional form, as {@link isRecordArgument} decides, and is checked under its old name.
 */
function optionOrPositional(
  value: unknown,
  key: string
): { readonly value: unknown; readonly field: string } {
  if (isRecordArgument(value)) {
    return { value: (value as Record<string, unknown>)[key], field: `options.${key}` };
  }
  return { value, field: key };
}

/** Options for `Build.mathBar`. */
export interface MathBarOptions {
  /** Where the bar is drawn. Default: `"top"`. */
  readonly position?: "top" | "bottom";
}

/** Create a math bar (overline or underline): `mathBar(content, { position })`. */
export function mathBar(content: readonly MathContent[], options?: MathBarOptions): MathContent;
/** @deprecated Pass the position in options: `mathBar(content, { position })`. Removed in 1.0. */
export function mathBar(
  content: readonly MathContent[],
  position: "top" | "bottom" | undefined
): MathContent;
export function mathBar(
  content: readonly MathContent[],
  positionOrOptions?: "top" | "bottom" | MathBarOptions
): MathContent {
  const where = "Build.mathBar";
  checkMath(where, { content });
  const position = optionOrPositional(positionOrOptions, "position");
  checkMathPosition(position.value, position.field, where);
  return {
    type: "mathBar",
    content,
    position: (position.value as "top" | "bottom" | undefined) ?? "top"
  };
}

/** Create a math box. */
export function mathBox(content: readonly MathContent[]): MathContent {
  checkMath("Build.mathBox", { content });
  return { type: "mathBox", content };
}

/** Create a math equation array. */
export function mathEquationArray(rows: readonly (readonly MathContent[])[]): MathContent {
  checkMathRows(rows, 1, "rows", "Build.mathEquationArray");
  return { type: "mathEquationArray", rows };
}

// =============================================================================
// Symbol / Image / Shape / Chart / SDT / CheckBox Builders
// =============================================================================

/** Create a symbol run. */
export function symbol(font: string, char: string, properties?: RunProperties): Run {
  checkString(font, "font", "Build.symbol");
  checkString(char, "char", "Build.symbol");
  checkRunProperties(properties, "properties", "Build.symbol");
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
  const where = "Build.floatingImage";
  checkRecord(options, "options", where);
  assertDrawingNumbers(where, options, under("options"));
  checkString(options.rId, "options.rId", where);
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
  assertDrawingNumbers("Build.drawingShape", options);
  assertDrawingNumber(options.outlineWidth, "outlineWidth", "Build.drawingShape", {
    sign: "nonNegative",
    optional: true
  });
  checkRgbColor(options.fillColor, "fillColor", "Build.drawingShape", { optional: true });
  checkRgbColor(options.outlineColor, "outlineColor", "Build.drawingShape", { optional: true });
  checkNodeArray(options.textContent, "textContent", "Build.drawingShape", { optional: true });
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
  const sized = { sign: "nonNegative", optional: true } as const;
  const where = "Build.chart";
  checkRecord(options, "options", where);
  checkOneOf(options.type, CHART_TYPES, "options.type", where);
  assertDrawingNumber(options.width, "options.width", where, sized);
  assertDrawingNumber(options.height, "options.height", where, sized);
  checkChartSeries(options.series, "options.series", where);
  checkString(options.title, "options.title", where, { optional: true });
  checkChartLegend(options.legend, "options.legend", where);
  checkRgbColor(options.plotAreaColor, "options.plotAreaColor", where, { optional: true });
  checkRgbColor(options.chartAreaColor, "options.chartAreaColor", where, { optional: true });
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
  content: readonly (Paragraph | Run | Table | StructuredDocumentTag)[],
  properties?: SdtProperties
): StructuredDocumentTag {
  checkNodeArray(content, "content", "Build.structuredDocumentTag");
  checkSdtProperties(properties, "properties", "Build.structuredDocumentTag");
  return { type: "sdt", properties: properties ?? {}, content };
}

/** Create a checkbox. */
export function checkBox(options?: {
  checked?: boolean;
  checkedState?: { value: string; font?: string };
  uncheckedState?: { value: string; font?: string };
}): CheckBox {
  const where = "Build.checkbox";
  checkBooleans(options, ["checked"], "options", where);
  for (const key of ["checkedState", "uncheckedState"] as const) {
    const state = options?.[key];
    if (state !== undefined) {
      checkRecord(state, `options.${key}`, where);
      checkString(state.value, `options.${key}.value`, where);
      checkString(state.font, `options.${key}.font`, where, { optional: true });
    }
  }
  return {
    type: "checkBox",
    checked: options?.checked,
    checkedState: options?.checkedState,
    uncheckedState: options?.uncheckedState
  };
}
