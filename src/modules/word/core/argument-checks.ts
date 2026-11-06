/**
 * The checks the public builders and editors (`Build.*`, `Document.*`) run on
 * what they are given, so that a value with no valid form in the package is
 * refused by the call that received it — with a {@link DocxArgumentError}
 * naming the function and the field — rather than written into a part Word
 * refuses, or rejected later by `Document.toBuffer` with an error that names an XML
 * attribute instead of the argument.
 *
 * The rules follow the schema types the values are written as: half-points,
 * twips and eighths of a point are integers (`ST_HpsMeasure`,
 * `ST_TwipsMeasure`, `ST_SignedTwipsMeasure`, `ST_EighthPointMeasure`), the
 * unsigned ones not negative; a colour is `RRGGBB` or `auto` (`ST_HexColor`);
 * an enumeration is one of its members. `Units.*` already rounds to the
 * integers these want.
 *
 * Only builder input is checked here. A document that was read may carry
 * values a builder would not accept, and is written as it is.
 */

import { describeValue } from "@utils/describe-value";
import { isRecordArgument } from "@utils/object";
import type { SchemaMatches } from "@utils/schema";
import { setAsideValue, unrecognizedValue } from "@utils/unrecognized-values";
import { checkSchema, w } from "@word/core/option-schema";
import { DocxArgumentError } from "@word/errors";
import { HASH_ALGORITHM_NAMES } from "@word/security/hash-names";
import type {
  Alignment,
  AppProperties,
  BodyContent,
  ChartLegendPosition,
  CoreProperties,
  DocxOptions,
  Border,
  BorderStyle,
  ChartType,
  DocumentSettings,
  DocumentGridType,
  DropCapType,
  EastAsianCombineBrackets,
  EmphasisMarkType,
  EndnoteProperties,
  FontSpec,
  FootnoteProperties,
  FrameAnchorType,
  HighlightColor,
  LineSpacingRule,
  MathBar,
  MathFraction,
  MathGroupChar,
  MathLimit,
  NoteNumberFormat,
  NoteNumberRestart,
  PageOrientation,
  PageVerticalAlign,
  ParagraphFrame,
  PositionalTabAlignment,
  PositionalTabLeader,
  PositionalTabRelativeTo,
  ProtectionType,
  ParagraphProperties,
  RevisionInfo,
  RowHeightRule,
  RubyAlign,
  RubyProperties,
  RunProperties,
  SdtDateProperties,
  SdtListItem,
  SdtProperties,
  SectionBreakType,
  SectionProperties,
  Shading,
  ShadingType,
  StyleDef,
  StyleType,
  TableOfContents,
  TabStopLeader,
  TabStopType,
  TableCellProperties,
  TableLayout,
  TableProperties,
  TableRowProperties,
  TableWidth,
  TextDirection,
  TextEffect,
  UnderlineStyle,
  VerticalAlign,
  VerticalCellAlign,
  VerticalMerge,
  Watermark,
  FloatingImage,
  HorizontalPositionRelative,
  VerticalPositionRelative,
  WrapStyle,
  WrapTextSide
} from "@word/types";

// -----------------------------------------------------------------------------
// Enumerations
// -----------------------------------------------------------------------------

/**
 * A runtime set of every member of the string union `T`. The type parameter
 * makes a missing member a compile error, so the set cannot fall behind the
 * union it checks.
 */
function members<T extends string>() {
  return <const V extends readonly T[]>(
    values: V &
      ([Exclude<T, V[number]>] extends [never] ? unknown : { missing: Exclude<T, V[number]> })
  ): ReadonlySet<string> => new Set<string>(values);
}

const UNDERLINE_STYLES = members<UnderlineStyle>()([
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble",
  "none"
]);

const HIGHLIGHT_COLORS = members<HighlightColor>()([
  "black",
  "blue",
  "cyan",
  "darkBlue",
  "darkCyan",
  "darkGray",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "green",
  "lightGray",
  "magenta",
  "none",
  "red",
  "white",
  "yellow"
]);

const VERTICAL_ALIGNS = members<VerticalAlign>()(["superscript", "subscript", "baseline"]);

const TEXT_EFFECTS = members<TextEffect>()([
  "blinkBackground",
  "lights",
  "antsBlack",
  "antsRed",
  "shimmer",
  "sparkle",
  "none"
]);

const EMPHASIS_MARKS = members<EmphasisMarkType>()(["dot", "comma", "circle", "underDot", "none"]);

const COMBINE_BRACKETS = members<EastAsianCombineBrackets>()([
  "none",
  "round",
  "square",
  "angle",
  "curly"
]);

const FONT_HINTS = members<NonNullable<FontSpec["hint"]>>()(["default", "eastAsia", "cs"]);

const BORDER_STYLES = members<BorderStyle>()([
  "single",
  "double",
  "thick",
  "dotted",
  "dashed",
  "dotDash",
  "dotDotDash",
  "triple",
  "thinThickSmallGap",
  "thickThinSmallGap",
  "thinThickThinSmallGap",
  "thinThickMediumGap",
  "thickThinMediumGap",
  "thinThickThinMediumGap",
  "thinThickLargeGap",
  "thickThinLargeGap",
  "thinThickThinLargeGap",
  "wave",
  "doubleWave",
  "dashSmallGap",
  "dashDotStroked",
  "threeDEmboss",
  "threeDEngrave",
  "outset",
  "inset",
  "none",
  "nil"
]);

const SHADING_TYPES = members<ShadingType>()([
  "clear",
  "solid",
  "horzStripe",
  "vertStripe",
  "reverseDiagStripe",
  "diagStripe",
  "horzCross",
  "diagCross",
  "thinHorzStripe",
  "thinVertStripe",
  "thinReverseDiagStripe",
  "thinDiagStripe",
  "thinHorzCross",
  "thinDiagCross",
  "pct5",
  "pct10",
  "pct12",
  "pct15",
  "pct20",
  "pct25",
  "pct30",
  "pct35",
  "pct37",
  "pct40",
  "pct45",
  "pct50",
  "pct55",
  "pct60",
  "pct62",
  "pct65",
  "pct70",
  "pct75",
  "pct80",
  "pct85",
  "pct87",
  "pct90",
  "pct95",
  "nil"
]);

const ALIGNMENTS = members<Alignment>()([
  "start",
  "center",
  "end",
  "both",
  "mediumKashida",
  "distribute",
  "numTab",
  "highKashida",
  "lowKashida",
  "thaiDistribute",
  "left",
  "right"
]);

const LINE_SPACING_RULES = members<LineSpacingRule>()(["auto", "exact", "atLeast"]);

const TAB_STOP_TYPES = members<TabStopType>()([
  "left",
  "center",
  "right",
  "decimal",
  "bar",
  "clear",
  "start",
  "end",
  "num"
]);

const TAB_STOP_LEADERS = members<TabStopLeader>()([
  "dot",
  "hyphen",
  "underscore",
  "none",
  "heavy",
  "middleDot"
]);

const TEXT_DIRECTIONS = members<TextDirection>()([
  "lrTb",
  "tbRl",
  "btLr",
  "lrTbV",
  "tbRlV",
  "tbLrV"
]);

const TEXT_ALIGNMENTS = members<NonNullable<ParagraphProperties["textAlignment"]>>()([
  "auto",
  "baseline",
  "bottom",
  "center",
  "top"
]);

const DROP_CAPS = members<DropCapType>()(["drop", "margin", "none"]);

const FRAME_WRAPS = members<NonNullable<ParagraphFrame["wrap"]>>()([
  "auto",
  "around",
  "none",
  "notBeside",
  "through",
  "tight"
]);

const FRAME_ANCHORS = members<FrameAnchorType>()(["margin", "page", "text"]);

const FRAME_X_ALIGNS = members<NonNullable<ParagraphFrame["xAlign"]>>()([
  "left",
  "center",
  "right",
  "inside",
  "outside"
]);

const FRAME_Y_ALIGNS = members<NonNullable<ParagraphFrame["yAlign"]>>()([
  "top",
  "center",
  "bottom",
  "inside",
  "outside",
  "inline"
]);

const TABLE_WIDTH_TYPES = members<TableWidth["type"]>()(["auto", "dxa", "pct", "nil"]);

const TABLE_LAYOUTS = members<TableLayout>()(["fixed", "autofit"]);

const ROW_HEIGHT_RULES = members<RowHeightRule>()(["auto", "atLeast", "exact"]);

const VERTICAL_MERGES = members<VerticalMerge>()(["restart", "continue"]);

const VERTICAL_CELL_ALIGNS = members<VerticalCellAlign>()(["top", "center", "bottom"]);

const SECTION_BREAKS = members<SectionBreakType>()([
  "nextPage",
  "continuous",
  "evenPage",
  "oddPage",
  "nextColumn"
]);

const ORIENTATIONS = members<PageOrientation>()(["portrait", "landscape"]);

const PAGE_VERTICAL_ALIGNS = members<PageVerticalAlign>()(["top", "center", "both", "bottom"]);

const DOC_GRID_TYPES = members<DocumentGridType>()([
  "default",
  "lines",
  "linesAndChars",
  "snapToChars"
]);

type LineNumberRestart = NonNullable<NonNullable<SectionProperties["lineNumbers"]>["restart"]>;
const LINE_NUMBER_RESTARTS = members<LineNumberRestart>()(["newPage", "newSection", "continuous"]);

type PageBorderDisplay = NonNullable<NonNullable<SectionProperties["pageBorders"]>["display"]>;
const PAGE_BORDER_DISPLAYS = members<PageBorderDisplay>()([
  "allPages",
  "firstPage",
  "notFirstPage"
]);

const PAGE_BORDER_OFFSETS = members<"page" | "text">()(["page", "text"]);

const PAGE_BORDER_Z_ORDERS = members<"front" | "back">()(["front", "back"]);

const NOTE_NUMBER_FORMATS = members<NoteNumberFormat>()([
  "decimal",
  "upperRoman",
  "lowerRoman",
  "upperLetter",
  "lowerLetter",
  "chicago",
  "bullet"
]);

const NOTE_NUMBER_RESTARTS = members<NoteNumberRestart>()(["continuous", "eachSect", "eachPage"]);

const FOOTNOTE_POSITIONS = members<NonNullable<FootnoteProperties["position"]>>()([
  "pageBottom",
  "beneathText"
]);

const ENDNOTE_POSITIONS = members<NonNullable<EndnoteProperties["position"]>>()([
  "sectEnd",
  "docEnd"
]);

// -----------------------------------------------------------------------------
// Primitive checks
// -----------------------------------------------------------------------------

/**
 * A field's path inside an object a function was given, mapped to its path from
 * the parameter: `under("options")` turns `wrap.style` into `options.wrap.style`,
 * as every argument error names a field (see AGENTS.md, "An argument error says
 * where").
 */
export type FieldName = (path: string) => string;

/** A {@link FieldName} for fields that are themselves parameters. */
export const sameField: FieldName = path => path;

/** A {@link FieldName} for the fields of the parameter `parameter`. */
export function under(parameter: string): FieldName {
  return path => `${parameter}.${path}`;
}

/**
 * The `where` of a check that cannot know its public caller: the writer's
 * validation of a document model, which every writing entry point (`Document.toBuffer`,
 * `Document.writeFile`, `Package.*`, the streaming writer) shares. Its message states
 * the problem without a prefix, and `field` (the model path) plus the stack
 * locate it — see AGENTS.md, "An argument error says where".
 */
export const MODEL_CHECK = "";

/** `"<where>: "`, or nothing for a {@link MODEL_CHECK}. */
export function messagePrefix(where: string): string {
  return where === MODEL_CHECK ? "" : `${where}: `;
}

/**
 * Throw the error every check reports: `<where>: <field> must be <expected>, got <value>`
 * (no `<where>: ` for a {@link MODEL_CHECK}).
 */
export function invalidArgument(
  where: string,
  field: string,
  expected: string,
  value: unknown
): never {
  throw new DocxArgumentError(
    field,
    `${messagePrefix(where)}${field} must be ${expected}, got ${describeValue(value)}`
  );
}

/** A short, readable rendering of a rejected value — shared with every module (`@utils/describe-value`). */
export { describeValue } from "@utils/describe-value";

interface Optional {
  /** `undefined` passes. */
  readonly optional?: boolean;
}

export function checkString(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if ((value === undefined && options.optional) || typeof value === "string") {
    return;
  }
  invalidArgument(where, field, "a string", value);
}

/**
 * `xsd:dateTime` lexical form — `YYYY-MM-DDThh:mm:ss`, optional fraction and
 * zone — which is what WordprocessingML's `ST_DateTime` (`w:date`) is.
 */
const XSD_DATE_TIME =
  /^-?\d{4,}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])T([01]\d|2[0-3]):[0-5]\d:[0-5]\d(\.\d+)?(Z|[+-](0\d|1[0-4]):[0-5]\d)?$/;

/** A string Word reads as a date: an `xsd:dateTime`, such as `new Date().toISOString()` gives. */
export function checkDateTime(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  if (typeof value !== "string" || !XSD_DATE_TIME.test(value) || Number.isNaN(Date.parse(value))) {
    invalidArgument(where, field, 'an xsd:dateTime such as "2024-01-31T09:30:00Z"', value);
  }
}

/**
 * The options every function that writes a package takes: `compressionLevel` a ZIP deflate
 * level (0–9) and `modTime`, when there is one, a valid `Date`. Checked here so a bad level is
 * the Word API's error naming the option, not the archive's naming its own parameter.
 */
export const packageOptionsShape = {
  compressionLevel: w.integer({ min: 0, max: 9 }).optional(),
  modTime: w.validDate().optional()
};

const docxOptionsSchema = w.record({ compressionLevel: packageOptionsShape.compressionLevel });
const _docxOptionsMatch: SchemaMatches<typeof docxOptionsSchema, DocxOptions> = true;
void _docxOptionsMatch;

const packageOptionsSchema = w.record(packageOptionsShape).optional();

/** The options of a function that writes a package: see {@link packageOptionsShape}. */
export function checkPackageOptions(value: unknown, field: string, where: string): void {
  checkSchema(packageOptionsSchema, value, where, field);
}

/**
 * A size limit: a number ≥ 0, `Infinity` for none. `NaN` makes every
 * comparison false, which would switch the limit off.
 */
export function checkLimit(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  if (typeof value !== "number" || !(value >= 0)) {
    invalidArgument(where, field, "a number ≥ 0 (Infinity for no limit)", value);
  }
}

/** A `Date` that holds a time (`new Date("x")` does not). */
export function checkValidDate(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    invalidArgument(where, field, "a valid Date", value);
  }
}

export function checkBoolean(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if ((value === undefined && options.optional) || typeof value === "boolean") {
    return;
  }
  invalidArgument(where, field, "a boolean", value);
}

/** The handles Word's functions take, each told apart by the list it always holds. */
const HANDLES = {
  document: { list: "body", expected: "a DocxDocument (from Document.create or Document.read)" },
  glossary: {
    list: "blocks",
    expected: "a GlossaryDocument (from Glossary.create or doc.glossary)"
  },
  styleMap: { list: "rules", expected: "a StyleMap (from Styles.parse or Styles.create)" }
} as const;

/**
 * The handle a public function was given — `doc`, a glossary, a style map —
 * checked as it crosses the boundary: an object holding the handle's list.
 * One `typeof` and one `Array.isArray`, so it costs nothing measurable even on
 * a function called per paragraph, and it turns the `TypeError` an untyped
 * caller would meet somewhere inside into an error naming the argument.
 */
export function checkHandle(
  value: unknown,
  kind: keyof typeof HANDLES,
  field: string,
  where: string
): void {
  const handle = HANDLES[kind];
  if (
    typeof value !== "object" ||
    value === null ||
    !Array.isArray((value as Record<string, unknown>)[handle.list])
  ) {
    invalidArgument(where, field, handle.expected, value);
  }
}

/**
 * Bytes as every Word entry point takes them: any `ArrayBufferView` — a
 * `Uint8Array` or Node `Buffer` (returned as it is), a `DataView` or another
 * typed array (read as the bytes it views, without a copy) — or an
 * `ArrayBuffer`, as `Workbook.read` and the PDF readers take them. Anything
 * else is refused.
 *
 * `ArrayBuffer.isView` alone is not a check: it admits a `DataView`, which has
 * no `slice` or index, and a wider typed array, whose elements are not bytes.
 */
export function checkBytes(value: unknown, field: string, where: string): Uint8Array {
  if (value instanceof Uint8Array) {
    return value;
  }
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  // Recognised by its tag, which an `ArrayBuffer` from another realm also carries.
  if (Object.prototype.toString.call(value) === "[object ArrayBuffer]") {
    return new Uint8Array(value as ArrayBuffer);
  }
  return invalidArgument(
    where,
    field,
    "a Uint8Array, another ArrayBufferView or an ArrayBuffer",
    value
  );
}

/**
 * Name → value data as the filling functions take it (`Document.mailMerge`,
 * `Template.fill`, `Transform.fillFormFields`, `Transform.resolveDataBindings`):
 * a plain object or a `Map` keyed by string. Returns its entries; a value is
 * named `<field>["<key>"]` whichever form it came in.
 */
export function checkNamedValues(
  value: unknown,
  field: string,
  where: string
): [string, unknown][] {
  if (value instanceof Map) {
    const entries: [string, unknown][] = [];
    for (const [key, v] of value) {
      if (typeof key !== "string") {
        invalidArgument(where, field, "a Map keyed by string", key);
      }
      entries.push([key, v]);
    }
    return entries;
  }
  if (!isRecordArgument(value)) {
    invalidArgument(where, field, "an object or a Map", value);
  }
  return Object.entries(value as Record<string, unknown>);
}

/** The field a {@link checkNamedValues} entry is reported under. */
export function namedValueField(field: string, key: string): string {
  return `${field}[${JSON.stringify(key)}]`;
}

/** A non-null object that is not an array, typed array or `Date`. */
export function checkRecord(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if ((value === undefined && options.optional) || isRecordArgument(value)) {
    return;
  }
  invalidArgument(where, field, "an object", value);
}

export function checkArray(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if ((value === undefined && options.optional) || Array.isArray(value)) {
    return;
  }
  invalidArgument(where, field, "an array", value);
}

/** Every element of an array is an object — a node built by `Build.*` or by hand. */
export function checkNodeArray(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  checkArray(value, field, where);
  (value as readonly unknown[]).forEach((node, i) => checkRecord(node, `${field}[${i}]`, where));
}

/** A node of the given `type`, e.g. a paragraph passed to `Document.addParagraph`. */
export function checkNode(value: unknown, type: string, field: string, where: string): void {
  if (isRecordArgument(value) && (value as { type?: unknown }).type === type) {
    return;
  }
  invalidArgument(where, field, `a ${type} node`, value);
}

interface IntegerOptions extends Optional {
  readonly min?: number;
  readonly max?: number;
}

/** A finite integer within `[min, max]`. */
export function checkInteger(
  value: unknown,
  field: string,
  where: string,
  options: IntegerOptions = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  const { min, max } = options;
  if (
    typeof value === "number" &&
    Number.isInteger(value) &&
    (min === undefined || value >= min) &&
    (max === undefined || value <= max)
  ) {
    return;
  }
  const range =
    min !== undefined && max !== undefined
      ? ` ${min}–${max}`
      : min !== undefined
        ? ` ≥ ${min}`
        : max !== undefined
          ? ` ≤ ${max}`
          : "";
  invalidArgument(where, field, `an integer${range}`, value);
}

/** A finite number: at least `min`, or above zero when `positive`. */
export function checkFinite(
  value: unknown,
  field: string,
  where: string,
  options: Optional & { readonly min?: number; readonly positive?: boolean } = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    (options.min === undefined || value >= options.min) &&
    (!options.positive || value > 0)
  ) {
    return;
  }
  invalidArgument(
    where,
    field,
    options.positive
      ? "a finite number > 0"
      : options.min === undefined
        ? "a finite number"
        : `a finite number ≥ ${options.min}`,
    value
  );
}

export function checkOneOf(
  value: unknown,
  allowed: ReadonlySet<string>,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (
    (value === undefined && options.optional) ||
    (typeof value === "string" && allowed.has(value))
  ) {
    return;
  }
  invalidArgument(
    where,
    field,
    `one of ${[...allowed].map(v => JSON.stringify(v)).join(", ")}`,
    value
  );
}

const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;
const HEX_BYTE = /^[0-9A-Fa-f]{2}$/;

/** `ST_HexColor`: six hex digits (`RRGGBB`, no `#`) or `"auto"`. */
export function checkHexColor(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (
    (value === undefined && options.optional) ||
    (typeof value === "string" && (value === "auto" || HEX_COLOR.test(value)))
  ) {
    return;
  }
  invalidArgument(where, field, 'a hex color "RRGGBB" (no "#") or "auto"', value);
}

/** A DrawingML or VML colour, `RRGGBB` (no `#`); unlike `ST_HexColor` it has no `"auto"`. */
export function checkRgbColor(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (
    (value === undefined && options.optional) ||
    (typeof value === "string" && HEX_COLOR.test(value))
  ) {
    return;
  }
  invalidArgument(where, field, 'a hex color "RRGGBB" (no "#")', value);
}

/** A theme tint or shade (`ST_UcharHexNumber`): two hex digits. Absent is accepted. */
export function checkHexByte(value: unknown, field: string, where: string): void {
  if (value === undefined || (typeof value === "string" && HEX_BYTE.test(value))) {
    return;
  }
  invalidArgument(where, field, 'two hex digits, e.g. "99"', value);
}

/** A heading level as `Build.heading` and `Document.addHeading` take it. */
export function checkHeadingLevel(value: unknown, field: string, where: string): void {
  checkInteger(value, field, where, { min: 1, max: 9 });
}

/** A list level (`w:ilvl`, 0–8). */
export function checkListLevel(value: unknown, field: string, where: string): void {
  checkInteger(value, field, where, { min: 0, max: 8 });
}

/** An id of a bookmark, comment or revision (`ST_DecimalNumber`, a signed integer). */
export function checkId(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  checkInteger(value, field, where, options);
}

// -----------------------------------------------------------------------------
// Property records
// -----------------------------------------------------------------------------

const optional = { optional: true } as const;
const unsigned = { optional: true, min: 0 } as const;

function twips(value: unknown, field: string, where: string): void {
  checkInteger(value, field, where, optional);
}

function unsignedTwips(value: unknown, field: string, where: string): void {
  checkInteger(value, field, where, unsigned);
}

export function checkRevision(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const rev = value as RevisionInfo;
  checkString(rev.author, `${field}.author`, where);
  checkId(rev.id, `${field}.id`, where);
  checkDateTime(rev.date, `${field}.date`, where, optional);
}

export function checkBorderStyle(value: unknown, field: string, where: string): void {
  checkOneOf(value, BORDER_STYLES, field, where);
}

export function checkBorder(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const b = value as Border;
  checkOneOf(b.style, BORDER_STYLES, `${field}.style`, where);
  checkInteger(b.size, `${field}.size`, where, unsigned);
  checkFinite(b.space, `${field}.space`, where, optional);
  checkHexColor(b.color, `${field}.color`, where, optional);
  checkString(b.themeColor, `${field}.themeColor`, where, optional);
  checkString(b.art, `${field}.art`, where, optional);
}

function checkBorderSet(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  for (const [side, border] of Object.entries(value as Record<string, unknown>)) {
    if (border !== undefined) {
      checkBorder(border, `${field}.${side}`, where);
    }
  }
}

export function checkShading(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const s = value as Shading;
  checkOneOf(s.pattern, SHADING_TYPES, `${field}.pattern`, where, optional);
  checkHexColor(s.color, `${field}.color`, where, optional);
  checkHexColor(s.fill, `${field}.fill`, where);
}

/**
 * Check run (character) properties. `field` is the path of the record in the
 * caller's arguments, e.g. `properties` or `properties.run`.
 */
export function checkRunProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const r = value as RunProperties;
  const at = (name: string): string => `${field}.${name}`;

  if (r.font !== undefined && typeof r.font !== "string") {
    checkRecord(r.font, at("font"), where);
    const f = r.font;
    for (const key of [
      "ascii",
      "hAnsi",
      "eastAsia",
      "cs",
      "asciiTheme",
      "hAnsiTheme",
      "eastAsiaTheme",
      "cstheme"
    ] as const) {
      checkString(f[key], at(`font.${key}`), where, optional);
    }
    checkOneOf(f.hint, FONT_HINTS, at("font.hint"), where, optional);
  }
  checkInteger(r.size, at("size"), where, unsigned);
  checkInteger(r.sizeCs, at("sizeCs"), where, unsigned);
  checkInteger(r.kern, at("kern"), where, unsigned);
  checkInteger(r.position, at("position"), where, optional);
  checkInteger(r.spacing, at("spacing"), where, optional);
  // ST_TextScale: a percentage, 1–600.
  checkInteger(r.scale, at("scale"), where, { optional: true, min: 1, max: 600 });

  if (typeof r.color === "string" || r.color === undefined) {
    checkHexColor(r.color, at("color"), where, optional);
  } else {
    checkRecord(r.color, at("color"), where);
    checkHexColor(r.color.val, at("color.val"), where);
    checkString(r.color.themeColor, at("color.themeColor"), where, optional);
    checkHexByte(r.color.themeTint, at("color.themeTint"), where);
    checkHexByte(r.color.themeShade, at("color.themeShade"), where);
  }
  checkOneOf(r.highlight, HIGHLIGHT_COLORS, at("highlight"), where, optional);
  if (r.underline !== undefined && typeof r.underline !== "boolean") {
    if (isRecordArgument(r.underline)) {
      checkOneOf(r.underline.style, UNDERLINE_STYLES, at("underline.style"), where);
      checkHexColor(r.underline.color, at("underline.color"), where, optional);
    } else {
      checkOneOf(r.underline, UNDERLINE_STYLES, at("underline"), where);
    }
  }
  if (r.shading !== undefined) {
    checkShading(r.shading, at("shading"), where);
  }
  if (r.border !== undefined) {
    checkBorder(r.border, at("border"), where);
  }
  checkOneOf(r.vertAlign, VERTICAL_ALIGNS, at("vertAlign"), where, optional);
  checkOneOf(r.effect, TEXT_EFFECTS, at("effect"), where, optional);
  checkOneOf(r.emphasisMark, EMPHASIS_MARKS, at("emphasisMark"), where, optional);
  checkString(r.style, at("style"), where, optional);
  if (r.language !== undefined) {
    checkRecord(r.language, at("language"), where);
    for (const key of ["val", "eastAsia", "bidi"] as const) {
      checkString(r.language[key], at(`language.${key}`), where, optional);
    }
  }
  if (r.eastAsianLayout !== undefined) {
    checkRecord(r.eastAsianLayout, at("eastAsianLayout"), where);
    checkInteger(r.eastAsianLayout.id, at("eastAsianLayout.id"), where, unsigned);
    checkOneOf(
      r.eastAsianLayout.combineBrackets,
      COMBINE_BRACKETS,
      at("eastAsianLayout.combineBrackets"),
      where,
      optional
    );
  }
  if (r.fitText !== undefined) {
    checkRecord(r.fitText, at("fitText"), where);
    checkInteger(r.fitText.val, at("fitText.val"), where, { min: 0 });
    checkInteger(r.fitText.id, at("fitText.id"), where, optional);
  }
  if (r.propertyChange !== undefined) {
    checkRecord(r.propertyChange, at("propertyChange"), where);
    checkRevision(r.propertyChange.revision, at("propertyChange.revision"), where);
    checkRunProperties(
      r.propertyChange.previousProperties,
      at("propertyChange.previousProperties"),
      where
    );
  }
}

function checkFrame(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const f = value as ParagraphFrame;
  const at = (name: string): string => `${field}.${name}`;
  checkOneOf(f.dropCap, DROP_CAPS, at("dropCap"), where, optional);
  checkInteger(f.lines, at("lines"), where, unsigned);
  unsignedTwips(f.width, at("width"), where);
  unsignedTwips(f.height, at("height"), where);
  unsignedTwips(f.hSpace, at("hSpace"), where);
  unsignedTwips(f.vSpace, at("vSpace"), where);
  twips(f.x, at("x"), where);
  twips(f.y, at("y"), where);
  checkOneOf(f.wrap, FRAME_WRAPS, at("wrap"), where, optional);
  checkOneOf(f.hAnchor, FRAME_ANCHORS, at("hAnchor"), where, optional);
  checkOneOf(f.vAnchor, FRAME_ANCHORS, at("vAnchor"), where, optional);
  checkOneOf(f.xAlign, FRAME_X_ALIGNS, at("xAlign"), where, optional);
  checkOneOf(f.yAlign, FRAME_Y_ALIGNS, at("yAlign"), where, optional);
}

/**
 * Check paragraph properties; `run` — the text form's run properties
 * (`textParagraph`, `Document.addParagraph`) — is checked when present.
 */
export function checkParagraphProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const p = value as ParagraphProperties & { run?: unknown };
  const at = (name: string): string => `${field}.${name}`;

  checkString(p.style, at("style"), where, optional);
  checkOneOf(p.alignment, ALIGNMENTS, at("alignment"), where, optional);
  if (p.indent !== undefined) {
    checkRecord(p.indent, at("indent"), where);
    for (const key of ["left", "right", "start", "end"] as const) {
      twips(p.indent[key], at(`indent.${key}`), where);
    }
    unsignedTwips(p.indent.hanging, at("indent.hanging"), where);
    unsignedTwips(p.indent.firstLine, at("indent.firstLine"), where);
  }
  if (p.spacing !== undefined) {
    checkRecord(p.spacing, at("spacing"), where);
    unsignedTwips(p.spacing.before, at("spacing.before"), where);
    unsignedTwips(p.spacing.after, at("spacing.after"), where);
    twips(p.spacing.line, at("spacing.line"), where);
    checkOneOf(p.spacing.lineRule, LINE_SPACING_RULES, at("spacing.lineRule"), where, optional);
  }
  if (p.numbering !== undefined) {
    checkRecord(p.numbering, at("numbering"), where);
    checkInteger(p.numbering.numId, at("numbering.numId"), where, { min: 0 });
    // The writer defaults an absent level to 0.
    if (p.numbering.level !== undefined) {
      checkListLevel(p.numbering.level, at("numbering.level"), where);
    }
  }
  checkBorderSet(p.borders, at("borders"), where);
  if (p.shading !== undefined) {
    checkShading(p.shading, at("shading"), where);
  }
  if (p.tabs !== undefined) {
    checkArray(p.tabs, at("tabs"), where);
    p.tabs.forEach((tab, i) => {
      const t = at(`tabs[${i}]`);
      checkRecord(tab, t, where);
      checkOneOf(tab.type, TAB_STOP_TYPES, `${t}.type`, where);
      checkInteger(tab.position, `${t}.position`, where);
      checkOneOf(tab.leader, TAB_STOP_LEADERS, `${t}.leader`, where, optional);
    });
  }
  checkInteger(p.outlineLevel, at("outlineLevel"), where, { optional: true, min: 0, max: 9 });
  if (p.frame !== undefined) {
    checkFrame(p.frame, at("frame"), where);
  }
  checkOneOf(p.textAlignment, TEXT_ALIGNMENTS, at("textAlignment"), where, optional);
  checkOneOf(p.textDirection, TEXT_DIRECTIONS, at("textDirection"), where, optional);
  checkRunProperties(p.markRunProperties, at("markRunProperties"), where);
  checkRunProperties(p.run, at("run"), where);
  checkSectionProperties(p.sectionProperties, at("sectionProperties"), where);
  if (p.paragraphInsertion !== undefined) {
    checkRevision(p.paragraphInsertion, at("paragraphInsertion"), where);
  }
  if (p.paragraphDeletion !== undefined) {
    checkRevision(p.paragraphDeletion, at("paragraphDeletion"), where);
  }
}

export function checkTableWidth(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  checkRecord(value, field, where);
  const w = value as TableWidth;
  checkOneOf(w.type, TABLE_WIDTH_TYPES, `${field}.type`, where);
  checkInteger(w.value, `${field}.value`, where);
}

function checkCellMargins(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  for (const [side, width] of Object.entries(value as Record<string, unknown>)) {
    checkTableWidth(width, `${field}.${side}`, where, optional);
  }
}

export function checkTableProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const t = value as TableProperties;
  const at = (name: string): string => `${field}.${name}`;
  checkString(t.style, at("style"), where, optional);
  checkTableWidth(t.width, at("width"), where, optional);
  checkBorderSet(t.borders, at("borders"), where);
  checkOneOf(t.layout, TABLE_LAYOUTS, at("layout"), where, optional);
  checkCellMargins(t.cellMargins, at("cellMargins"), where);
  checkTableAlignment(t.alignment, at("alignment"), where);
  twips(t.indent, at("indent"), where);
  checkInteger(t.rowBandSize, at("rowBandSize"), where, unsigned);
  checkInteger(t.colBandSize, at("colBandSize"), where, unsigned);
  checkTableWidth(t.cellSpacing, at("cellSpacing"), where, optional);
  if (t.shading !== undefined) {
    checkShading(t.shading, at("shading"), where);
  }
  checkString(t.caption, at("caption"), where, optional);
  checkString(t.description, at("description"), where, optional);
}

export function checkTableRowProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const r = value as TableRowProperties;
  const at = (name: string): string => `${field}.${name}`;
  if (r.height !== undefined) {
    checkRecord(r.height, at("height"), where);
    checkInteger(r.height.value, at("height.value"), where, { min: 0 });
    checkOneOf(r.height.rule, ROW_HEIGHT_RULES, at("height.rule"), where, optional);
  }
  checkTableWidth(r.cellSpacing, at("cellSpacing"), where, optional);
  checkInteger(r.gridBefore, at("gridBefore"), where, unsigned);
  checkInteger(r.gridAfter, at("gridAfter"), where, unsigned);
  checkTableWidth(r.widthBefore, at("widthBefore"), where, optional);
  checkTableWidth(r.widthAfter, at("widthAfter"), where, optional);
  checkTableProperties(r.tblPrEx, at("tblPrEx"), where);
}

export function checkTableCellProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const c = value as TableCellProperties;
  const at = (name: string): string => `${field}.${name}`;
  checkTableWidth(c.width, at("width"), where, optional);
  checkInteger(c.gridSpan, at("gridSpan"), where, { optional: true, min: 1 });
  checkInteger(c.rowSpan, at("rowSpan"), where, { optional: true, min: 1 });
  checkOneOf(c.verticalMerge, VERTICAL_MERGES, at("verticalMerge"), where, optional);
  checkBorderSet(c.borders, at("borders"), where);
  if (c.shading !== undefined) {
    checkShading(c.shading, at("shading"), where);
  }
  checkOneOf(c.verticalAlign, VERTICAL_CELL_ALIGNS, at("verticalAlign"), where, optional);
  checkOneOf(c.textDirection, TEXT_DIRECTIONS, at("textDirection"), where, optional);
  checkCellMargins(c.margins, at("margins"), where);
}

export function checkSectionProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const s = value as SectionProperties;
  const at = (name: string): string => `${field}.${name}`;
  if (s.pageSize !== undefined) {
    checkRecord(s.pageSize, at("pageSize"), where);
    // An absent dimension is written as the default page's.
    unsignedTwips(s.pageSize.width, at("pageSize.width"), where);
    unsignedTwips(s.pageSize.height, at("pageSize.height"), where);
    checkOneOf(s.pageSize.orientation, ORIENTATIONS, at("pageSize.orientation"), where, optional);
  }
  if (s.margins !== undefined) {
    checkRecord(s.margins, at("margins"), where);
    // Top and bottom are signed: a negative one lets text overlap a header. An
    // absent margin is written as the default.
    twips(s.margins.top, at("margins.top"), where);
    twips(s.margins.bottom, at("margins.bottom"), where);
    unsignedTwips(s.margins.left, at("margins.left"), where);
    unsignedTwips(s.margins.right, at("margins.right"), where);
    unsignedTwips(s.margins.header, at("margins.header"), where);
    unsignedTwips(s.margins.footer, at("margins.footer"), where);
    unsignedTwips(s.margins.gutter, at("margins.gutter"), where);
  }
  checkOneOf(s.breakType, SECTION_BREAKS, at("breakType"), where, optional);
  if (s.columns !== undefined) {
    checkRecord(s.columns, at("columns"), where);
    checkInteger(s.columns.count, at("columns.count"), where, { optional: true, min: 1 });
    unsignedTwips(s.columns.space, at("columns.space"), where);
    if (s.columns.columns !== undefined) {
      checkArray(s.columns.columns, at("columns.columns"), where);
      s.columns.columns.forEach((col, i) => {
        const c = at(`columns.columns[${i}]`);
        checkRecord(col, c, where);
        checkInteger(col.width, `${c}.width`, where, { min: 0 });
        unsignedTwips(col.space, `${c}.space`, where);
      });
    }
  }
  if (s.pageNumbering !== undefined) {
    checkRecord(s.pageNumbering, at("pageNumbering"), where);
    checkInteger(s.pageNumbering.start, at("pageNumbering.start"), where, unsigned);
    checkString(s.pageNumbering.format, at("pageNumbering.format"), where, optional);
  }
  if (s.docGrid !== undefined) {
    checkRecord(s.docGrid, at("docGrid"), where);
    twips(s.docGrid.linePitch, at("docGrid.linePitch"), where);
    checkInteger(s.docGrid.charSpace, at("docGrid.charSpace"), where, optional);
    checkOneOf(s.docGrid.type, DOC_GRID_TYPES, at("docGrid.type"), where, optional);
  }
  if (s.lineNumbers !== undefined) {
    checkRecord(s.lineNumbers, at("lineNumbers"), where);
    checkInteger(s.lineNumbers.countBy, at("lineNumbers.countBy"), where, unsigned);
    checkInteger(s.lineNumbers.start, at("lineNumbers.start"), where, unsigned);
    unsignedTwips(s.lineNumbers.distance, at("lineNumbers.distance"), where);
    checkOneOf(
      s.lineNumbers.restart,
      LINE_NUMBER_RESTARTS,
      at("lineNumbers.restart"),
      where,
      optional
    );
  }
  if (s.pageBorders !== undefined) {
    const pb = s.pageBorders;
    checkRecord(pb, at("pageBorders"), where);
    for (const side of ["top", "left", "bottom", "right"] as const) {
      if (pb[side] !== undefined) {
        checkBorder(pb[side], at(`pageBorders.${side}`), where);
      }
    }
    checkOneOf(pb.display, PAGE_BORDER_DISPLAYS, at("pageBorders.display"), where, optional);
    checkOneOf(pb.offsetFrom, PAGE_BORDER_OFFSETS, at("pageBorders.offsetFrom"), where, optional);
    checkOneOf(pb.zOrder, PAGE_BORDER_Z_ORDERS, at("pageBorders.zOrder"), where, optional);
  }
  checkOneOf(s.verticalAlign, PAGE_VERTICAL_ALIGNS, at("verticalAlign"), where, optional);
  checkOneOf(s.textDirection, TEXT_DIRECTIONS, at("textDirection"), where, optional);
  for (const [key, positions] of [
    ["footnoteProperties", FOOTNOTE_POSITIONS],
    ["endnoteProperties", ENDNOTE_POSITIONS]
  ] as const) {
    const notes = s[key];
    if (notes !== undefined) {
      checkRecord(notes, at(key), where);
      checkOneOf(notes.numFmt, NOTE_NUMBER_FORMATS, at(`${key}.numFmt`), where, optional);
      checkInteger(notes.numStart, at(`${key}.numStart`), where, unsigned);
      checkOneOf(notes.numRestart, NOTE_NUMBER_RESTARTS, at(`${key}.numRestart`), where, optional);
      checkOneOf(notes.position, positions, at(`${key}.position`), where, optional);
    }
  }
}

export function checkUnderlineStyle(value: unknown, field: string, where: string): void {
  checkOneOf(value, UNDERLINE_STYLES, field, where, optional);
}

/** A watermark as `Document.setWatermark` takes it. */
export function checkWatermark(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const w = value as Watermark;
  const at = (name: string): string => `${field}.${name}`;
  checkOneOf(w.type, WATERMARK_TYPES, at("type"), where);
  // A size of zero draws nothing: an invisible watermark is refused, not written.
  const size = { optional: true, positive: true } as const;
  if (w.type === "text") {
    if (typeof w.text !== "string" || w.text.length === 0) {
      invalidArgument(where, at("text"), "a non-empty string", w.text);
    }
    checkString(w.font, at("font"), where, optional);
    checkFinite(w.fontSize, at("fontSize"), where, size);
    // Written as a VML fill colour, `#RRGGBB`, which has no "auto".
    checkRgbColor(w.color, at("color"), where, optional);
    checkFinite(w.rotation, at("rotation"), where, optional);
    checkFinite(w.widthPt, at("widthPt"), where, size);
    checkFinite(w.heightPt, at("heightPt"), where, size);
  } else {
    checkString(w.rId, at("rId"), where);
    checkFinite(w.scale, at("scale"), where, unsigned);
    checkFinite(w.widthPt, at("widthPt"), where, size);
    checkFinite(w.heightPt, at("heightPt"), where, size);
  }
}

const WATERMARK_TYPES = members<Watermark["type"]>()(["text", "image"]);

const ZOOM_TYPES = members<NonNullable<DocumentSettings["zoomType"]>>()([
  "none",
  "fullPage",
  "bestFit",
  "textFit"
]);
const CHARACTER_SPACING = members<NonNullable<DocumentSettings["characterSpacingControl"]>>()([
  "doNotCompress",
  "compressPunctuation",
  "compressPunctuationAndJapaneseKana"
]);
/** `ST_DocProtect`. */
const PROTECTION_EDITS = members<ProtectionType>()([
  "none",
  "readOnly",
  "comments",
  "trackedChanges",
  "forms"
]);
const COMPATIBILITY_MODES: ReadonlySet<unknown> = new Set([11, 12, 14, 15]);
const SETTINGS_FLAGS = [
  "evenAndOddHeaders",
  "trackRevisions",
  "autoHyphenation",
  "mirrorMargins",
  "gutterAtTop",
  "displayBackgroundShape",
  "updateFieldsOnOpen",
  "doNotTrackMoves",
  "doNotTrackFormatting",
  "saveSubsetFonts",
  "bordersDoNotSurroundHeader",
  "bordersDoNotSurroundFooter",
  "noPunctuationKerning",
  "doNotDemoteAsianTextFirstLine"
] as const satisfies readonly (keyof DocumentSettings)[];
const SETTINGS_TEXT = [
  "decimalSymbol",
  "listSeparator",
  "clickAndTypeStyle",
  "stylePaneFormatFilter",
  "stylePaneSortMethod",
  "mailMergeRawXml"
] as const satisfies readonly (keyof DocumentSettings)[];

/**
 * Document settings as `Document.setSettings` and the streaming writer take
 * them: each value one `word/settings.xml` can hold, in the vocabularies the
 * dedicated functions (`setCompatibilityMode`, `Security.protect`) use.
 */
export function checkDocumentSettings(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const s = value as DocumentSettings;
  const at = (name: string): string => `${field}.${name}`;
  // A percentage (`w:zoom/@w:percent`).
  checkInteger(s.zoom, at("zoom"), where, { optional: true, min: 1 });
  checkOneOf(s.zoomType, ZOOM_TYPES, at("zoomType"), where, optional);
  unsignedTwips(s.defaultTabStop, at("defaultTabStop"), where);
  checkOneOf(
    s.characterSpacingControl,
    CHARACTER_SPACING,
    at("characterSpacingControl"),
    where,
    optional
  );
  if (s.compatibilityMode !== undefined && !COMPATIBILITY_MODES.has(s.compatibilityMode)) {
    invalidArgument(where, at("compatibilityMode"), "one of 11, 12, 14, 15", s.compatibilityMode);
  }
  for (const key of SETTINGS_FLAGS) {
    checkBoolean(s[key], at(key), where, optional);
  }
  for (const key of SETTINGS_TEXT) {
    checkString(s[key], at(key), where, optional);
  }
  if (s.documentProtection !== undefined) {
    const p = s.documentProtection;
    const pAt = (name: string): string => at(`documentProtection.${name}`);
    checkRecord(p, at("documentProtection"), where);
    checkOneOf(p.type, PROTECTION_EDITS, pAt("type"), where, optional);
    checkOneOf(p.edit, PROTECTION_EDITS, pAt("edit"), where, optional);
    checkOneOf(p.hashAlgorithm, HASH_ALGORITHM_NAMES, pAt("hashAlgorithm"), where, optional);
    checkBoolean(p.enforcement, pAt("enforcement"), where, optional);
    checkBoolean(p.formatting, pAt("formatting"), where, optional);
    checkString(p.hashValue, pAt("hashValue"), where, optional);
    checkString(p.saltValue, pAt("saltValue"), where, optional);
    checkInteger(p.spinCount, pAt("spinCount"), where, { optional: true, min: 0 });
  }
  for (const key of ["hyphenation", "writeProtection", "rsids", "themeFontLang"] as const) {
    checkRecord(s[key], at(key), where, optional);
  }
}

const CORE_TEXT = [
  "title",
  "subject",
  "creator",
  "description",
  "keywords",
  "lastModifiedBy",
  "revision",
  "category"
] as const;

/** Core properties (`docProps/core.xml`): text as strings, dates as valid `Date`s. */
export function checkCoreProperties(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const p = value as CoreProperties;
  for (const key of CORE_TEXT) {
    checkString(p[key], `${field}.${key}`, where, optional);
  }
  checkValidDate(p.created, `${field}.created`, where, optional);
  checkValidDate(p.modified, `${field}.modified`, where, optional);
}

const APP_TEXT = ["application", "appVersion", "company", "manager"] as const;
const APP_COUNTS = ["pages", "words", "characters", "lines", "paragraphs"] as const;

/** Application properties (`docProps/app.xml`): text as strings, statistics as integers ≥ 0. */
export function checkAppProperties(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const p = value as AppProperties;
  for (const key of APP_TEXT) {
    checkString(p[key], `${field}.${key}`, where, optional);
  }
  for (const key of APP_COUNTS) {
    checkInteger(p[key], `${field}.${key}`, where, { optional: true, min: 0 });
  }
}

/** Every `BodyContent["type"]`: what the body, a cell or a content control holds. */
const BLOCK_TYPES = members<BodyContent["type"]>()([
  "paragraph",
  "table",
  "floatingImage",
  "tableOfContents",
  "math",
  "textBox",
  "checkBox",
  "drawingShape",
  "opaqueDrawing",
  "chart",
  "chartEx",
  "altChunk",
  "sdt"
]);

/** A body block: an object whose `type` is one of the `BodyContent` kinds. */
export function checkBlock(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  checkOneOf((value as { type?: unknown }).type, BLOCK_TYPES, `${field}.type`, where);
}

const HORIZONTAL_RELATIVE = members<HorizontalPositionRelative>()([
  "character",
  "column",
  "insideMargin",
  "leftMargin",
  "margin",
  "outsideMargin",
  "page",
  "rightMargin"
]);
const VERTICAL_RELATIVE = members<VerticalPositionRelative>()([
  "insideMargin",
  "line",
  "margin",
  "outsideMargin",
  "page",
  "paragraph",
  "topMargin",
  "bottomMargin"
]);
const HORIZONTAL_ALIGNS = members<
  NonNullable<NonNullable<FloatingImage["horizontalPosition"]>["align"]>
>()(["left", "center", "right", "inside", "outside"]);
const VERTICAL_ALIGNS_ANCHOR = members<
  NonNullable<NonNullable<FloatingImage["verticalPosition"]>["align"]>
>()(["top", "center", "bottom", "inside", "outside"]);
const WRAP_STYLES = members<WrapStyle>()(["square", "tight", "through", "topAndBottom", "none"]);
const WRAP_SIDES = members<WrapTextSide>()(["bothSides", "left", "right", "largest"]);
/** `ST_LegendPos`; the model's `"none"` is no `c:legend` at all. */
const LEGEND_POSITIONS = members<Exclude<ChartLegendPosition, "none">>()([
  "b",
  "tr",
  "l",
  "r",
  "t"
]);
const CHART_LEGENDS = withValues(LEGEND_POSITIONS, ["none"]);
/** `ST_DLblPos`, the values a chart part holds; the model names four of them (see the chart writer). */
const DATA_LABEL_POSITIONS: ReadonlySet<string> = new Set([
  "bestFit",
  "b",
  "ctr",
  "inBase",
  "inEnd",
  "l",
  "outEnd",
  "r",
  "t"
]);

/**
 * A drawing anchor's enumerated fields — position `relativeTo` / `align` and
 * wrap `style` / `side` — and `wrap` an object, as every anchored drawing
 * (floating image, shape, chart frame) takes them.
 */
export function checkAnchorFields(
  fields: {
    readonly horizontalPosition?: unknown;
    readonly verticalPosition?: unknown;
    readonly wrap?: unknown;
  },
  where: string,
  field: FieldName = sameField
): void {
  const h = fields.horizontalPosition as { relativeTo?: unknown; align?: unknown } | undefined;
  if (h !== undefined) {
    checkRecord(h, field("horizontalPosition"), where);
    checkOneOf(
      h.relativeTo,
      HORIZONTAL_RELATIVE,
      field("horizontalPosition.relativeTo"),
      where,
      optional
    );
    checkOneOf(h.align, HORIZONTAL_ALIGNS, field("horizontalPosition.align"), where, optional);
  }
  const v = fields.verticalPosition as { relativeTo?: unknown; align?: unknown } | undefined;
  if (v !== undefined) {
    checkRecord(v, field("verticalPosition"), where);
    checkOneOf(
      v.relativeTo,
      VERTICAL_RELATIVE,
      field("verticalPosition.relativeTo"),
      where,
      optional
    );
    checkOneOf(v.align, VERTICAL_ALIGNS_ANCHOR, field("verticalPosition.align"), where, optional);
  }
  const wrap = fields.wrap as { style?: unknown; side?: unknown } | undefined;
  if (wrap !== undefined) {
    checkRecord(wrap, field("wrap"), where);
    checkOneOf(wrap.style, WRAP_STYLES, field("wrap.style"), where);
    checkOneOf(wrap.side, WRAP_SIDES, field("wrap.side"), where, optional);
  }
}

/** A chart's `legend`: an `ST_LegendPos` value or `"none"`. */
export function checkChartLegend(value: unknown, field: string, where: string): void {
  checkOneOf(value, CHART_LEGENDS, field, where, optional);
}

/** Every {@link ChartType}. */
export const CHART_TYPES = members<ChartType>()([
  "bar",
  "barStacked",
  "barPercentStacked",
  "column",
  "columnStacked",
  "columnPercentStacked",
  "line",
  "lineStacked",
  "lineMarked",
  "pie",
  "pie3D",
  "doughnut",
  "area",
  "areaStacked",
  "scatter",
  "scatterSmooth",
  "radar",
  "radarFilled",
  "bubble",
  "stock",
  "surface",
  "surface3D",
  "surfaceWireframe",
  "surfaceWireframe3D"
]);

/** An array of finite numbers — chart values, written as `c:v` text. */
export function checkFiniteArray(value: unknown, field: string, where: string): void {
  checkArray(value, field, where);
  (value as readonly unknown[]).forEach((v, i) => checkFinite(v, `${field}[${i}]`, where));
}

/** An array of strings. */
export function checkStringArray(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  checkArray(value, field, where);
  (value as readonly unknown[]).forEach((v, i) => checkString(v, `${field}[${i}]`, where));
}

/** A chart's series as `Build.chart` takes them. */
export function checkChartSeries(value: unknown, field: string, where: string): void {
  checkNodeArray(value, field, where);
  (value as readonly unknown[]).forEach((item, i) => {
    const at = `${field}[${i}]`;
    const series = item as Record<string, unknown>;
    checkString(series.name, `${at}.name`, where);
    checkStringArray(series.categories, `${at}.categories`, where);
    checkFiniteArray(series.values, `${at}.values`, where);
    checkRgbColor(series.color, `${at}.color`, where, optional);
    if (series.pointColors !== undefined) {
      checkArray(series.pointColors, `${at}.pointColors`, where);
      // Entry j is point j's colour; a point with none is a hole or `undefined`.
      (series.pointColors as readonly unknown[]).forEach((c, j) =>
        checkRgbColor(c, `${at}.pointColors[${j}]`, where, optional)
      );
    }
    checkOneOf(series.chartType, CHART_TYPES, `${at}.chartType`, where, optional);
  });
}

const POSITIONAL_TAB_ALIGNMENTS = members<PositionalTabAlignment>()(["center", "left", "right"]);
const POSITIONAL_TAB_RELATIVE = members<PositionalTabRelativeTo>()(["indent", "margin"]);
const POSITIONAL_TAB_LEADERS = members<PositionalTabLeader>()([
  "dot",
  "hyphen",
  "middleDot",
  "none",
  "underscore"
]);

export function checkPositionalTab(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const t = value as {
    alignment?: unknown;
    relativeTo?: unknown;
    leader?: unknown;
  };
  checkOneOf(t.alignment, POSITIONAL_TAB_ALIGNMENTS, `${field}.alignment`, where);
  checkOneOf(t.relativeTo, POSITIONAL_TAB_RELATIVE, `${field}.relativeTo`, where);
  checkOneOf(t.leader, POSITIONAL_TAB_LEADERS, `${field}.leader`, where, optional);
}

const STYLE_TYPES = members<StyleType>()(["paragraph", "character", "table", "numbering"]);

/** A style definition as `Document.addStyle` takes it. */
export function checkStyleDef(value: unknown, field: string, where: string): void {
  checkRecord(value, field, where);
  const s = value as StyleDef;
  const at = (name: string): string => `${field}.${name}`;
  checkOneOf(s.type, STYLE_TYPES, at("type"), where);
  checkString(s.styleId, at("styleId"), where);
  if (s.styleId === "") {
    invalidArgument(where, at("styleId"), "a non-empty string", s.styleId);
  }
  checkString(s.name, at("name"), where);
  for (const key of ["basedOn", "next", "link"] as const) {
    checkString(s[key], at(key), where, optional);
  }
  checkInteger(s.uiPriority, at("uiPriority"), where, unsigned);
  checkInteger(s.outlineLevel, at("outlineLevel"), where, { optional: true, min: 0, max: 9 });
  checkParagraphProperties(s.paragraphProperties, at("paragraphProperties"), where);
  checkRunProperties(s.runProperties, at("runProperties"), where);
  checkTableProperties(s.tableProperties, at("tableProperties"), where);
}

// -----------------------------------------------------------------------------
// Content builders: math, ruby, fields, form fields, content controls
// -----------------------------------------------------------------------------

/**
 * `ST_Char`: one character (one code point), or — where `empty` allows it, as
 * a delimiter does to mean "none" — the empty string.
 */
export function checkChar(
  value: unknown,
  field: string,
  where: string,
  options: Optional & { readonly empty?: boolean } = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  if (typeof value === "string") {
    let count = 0;
    for (const _ of value) {
      count++;
    }
    if (count === 1 || (count === 0 && options.empty)) {
      return;
    }
  }
  invalidArgument(
    where,
    field,
    options.empty ? "a single character or the empty string" : "a single character",
    value
  );
}

/** A string of at most `max` UTF-16 units — the schema's `maxLength`. */
export function checkBoundedString(
  value: unknown,
  max: number,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (
    (value === undefined && options.optional) ||
    (typeof value === "string" && value.length <= max)
  ) {
    return;
  }
  invalidArgument(where, field, `a string of at most ${max} characters`, value);
}

/**
 * A name a field instruction carries bare — a bookmark, a sequence identifier:
 * non-empty, with no white space or quote, which would end it early.
 */
export function checkFieldName(value: unknown, field: string, where: string): void {
  if (typeof value === "string" && /^[^\s"\\]+$/.test(value)) {
    return;
  }
  invalidArgument(where, field, "a non-empty name with no spaces, quotes or backslashes", value);
}

/** A TOC heading range, `"1-3"`: two levels 1–9, the first not above the second. */
export function checkHeadingRange(
  value: unknown,
  field: string,
  where: string,
  options: Optional = {}
): void {
  if (value === undefined && options.optional) {
    return;
  }
  const m = typeof value === "string" ? /^([1-9])-([1-9])$/.exec(value) : null;
  if (m && Number(m[1]) <= Number(m[2])) {
    return;
  }
  invalidArgument(where, field, 'a heading range "N-M" with 1 ≤ N ≤ M ≤ 9', value);
}

const MATH_FRACTION_TYPES = members<NonNullable<MathFraction["fractionType"]>>()([
  "bar",
  "skw",
  "lin",
  "noBar"
]);
const MATH_POSITIONS = members<MathBar["position"]>()(["top", "bottom"]);
const MATH_VERTICAL_ALIGNS = members<NonNullable<MathGroupChar["verticalAlign"]>>()([
  "top",
  "center",
  "bottom"
]);
const MATH_LIMIT_TYPES = members<MathLimit["limitType"]>()(["upper", "lower"]);

export function checkMathFractionType(value: unknown, field: string, where: string): void {
  checkOneOf(value, MATH_FRACTION_TYPES, field, where, optional);
}

export function checkMathPosition(value: unknown, field: string, where: string): void {
  checkOneOf(value, MATH_POSITIONS, field, where, optional);
}

export function checkMathVerticalAlign(value: unknown, field: string, where: string): void {
  checkOneOf(value, MATH_VERTICAL_ALIGNS, field, where, optional);
}

export function checkMathLimitType(value: unknown, field: string, where: string): void {
  checkOneOf(value, MATH_LIMIT_TYPES, field, where);
}

/** Rows of cells of math content, as a matrix takes them. */
export function checkMathRows(value: unknown, depth: 1 | 2, field: string, where: string): void {
  checkArray(value, field, where);
  (value as unknown[]).forEach((row, i) => {
    if (depth === 1) {
      checkNodeArray(row, `${field}[${i}]`, where);
    } else {
      checkMathRows(row, 1, `${field}[${i}]`, where);
    }
  });
}

/** Every value of `options` named in `keys` is a boolean or absent. */
export function checkBooleans(
  options: unknown,
  keys: readonly string[],
  field: string,
  where: string
): void {
  if (options === undefined) {
    return;
  }
  checkRecord(options, field, where);
  for (const key of keys) {
    checkBoolean((options as Record<string, unknown>)[key], `${field}.${key}`, where, optional);
  }
}

const RUBY_ALIGNS = members<RubyAlign>()([
  "center",
  "distributeLetter",
  "distributeSpace",
  "left",
  "right",
  "rightVertical"
]);

export function checkRubyProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const r = value as RubyProperties;
  checkOneOf(r.align, RUBY_ALIGNS, `${field}.align`, where, optional);
  checkInteger(r.fontSize, `${field}.fontSize`, where, unsigned);
  checkInteger(r.baseFontSize, `${field}.baseFontSize`, where, unsigned);
  checkInteger(r.raise, `${field}.raise`, where, optional);
  checkString(r.language, `${field}.language`, where, optional);
}

const SDT_APPEARANCES = members<NonNullable<SdtProperties["appearance"]>>()([
  "boundingBox",
  "tags",
  "hidden"
]);
const SDT_DATE_STORAGE = members<NonNullable<SdtDateProperties["storeMappedDataAs"]>>()([
  "dateTime",
  "date",
  "text"
]);

function checkSdtListItems(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkArray(value, field, where);
  (value as unknown[]).forEach((item, i) => {
    checkRecord(item, `${field}[${i}]`, where);
    const it = item as SdtListItem;
    checkString(it.value, `${field}[${i}].value`, where);
    checkString(it.displayText, `${field}[${i}].displayText`, where, optional);
  });
}

export function checkSdtProperties(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const p = value as SdtProperties;
  const at = (name: string): string => `${field}.${name}`;
  checkId(p.id, at("id"), where, optional);
  for (const key of ["tag", "alias", "placeholder"] as const) {
    checkString(p[key], at(key), where, optional);
  }
  checkSdtListItems(p.dropdownList, at("dropdownList"), where);
  checkSdtListItems(p.comboBox, at("comboBox"), where);
  if (p.date !== undefined) {
    checkRecord(p.date, at("date"), where);
    for (const key of ["fullDate", "dateFormat", "lid"] as const) {
      checkString(p.date[key], at(`date.${key}`), where, optional);
    }
    checkOneOf(
      p.date.storeMappedDataAs,
      SDT_DATE_STORAGE,
      at("date.storeMappedDataAs"),
      where,
      optional
    );
  }
  checkOneOf(p.appearance, SDT_APPEARANCES, at("appearance"), where, optional);
  if (p.dataBinding !== undefined) {
    checkRecord(p.dataBinding, at("dataBinding"), where);
    checkString(p.dataBinding.xpath, at("dataBinding.xpath"), where);
    checkString(p.dataBinding.storeItemId, at("dataBinding.storeItemId"), where);
    checkString(p.dataBinding.prefixMappings, at("dataBinding.prefixMappings"), where, optional);
  }
  if (p.checkbox !== undefined) {
    checkRecord(p.checkbox, at("checkbox"), where);
    const c = p.checkbox;
    checkBoolean(c.checked, at("checkbox.checked"), where, optional);
    for (const key of ["checkedChar", "uncheckedChar"] as const) {
      // The hex code point (`"2612"`) or the glyph itself (`"☒"`).
      if (c[key] !== undefined && !(typeof c[key] === "string" && c[key] !== "")) {
        invalidArgument(where, at(`checkbox.${key}`), "a hex code point or a character", c[key]);
      }
    }
    checkString(c.checkedFont, at("checkbox.checkedFont"), where, optional);
    checkString(c.uncheckedFont, at("checkbox.uncheckedFont"), where, optional);
  }
  if (p.repeatingSection !== undefined) {
    checkRecord(p.repeatingSection, at("repeatingSection"), where);
    checkString(p.repeatingSection.title, at("repeatingSection.title"), where, optional);
    checkString(
      p.repeatingSection.sectionTitle,
      at("repeatingSection.sectionTitle"),
      where,
      optional
    );
  }
  if (p.docPart !== undefined) {
    checkRecord(p.docPart, at("docPart"), where);
    checkString(p.docPart.gallery, at("docPart.gallery"), where, optional);
    checkString(p.docPart.category, at("docPart.category"), where, optional);
  }
}

const TOC_FIELD_MODES = members<NonNullable<TableOfContents["field"]>>()(["entries", "none"]);

/** The options of a table of contents, as `Document.addTableOfContents` takes them. */
export function checkTableOfContentsOptions(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  checkRecord(value, field, where);
  const t = value as Partial<TableOfContents>;
  const at = (name: string): string => `${field}.${name}`;
  checkString(t.instruction, at("instruction"), where, optional);
  checkHeadingRange(t.headingStyleRange, at("headingStyleRange"), where, optional);
  checkBoolean(t.hyperlink, at("hyperlink"), where, optional);
  if (t.stylesWithLevels !== undefined) {
    checkArray(t.stylesWithLevels, at("stylesWithLevels"), where);
    t.stylesWithLevels.forEach((entry, i) => {
      const e = at(`stylesWithLevels[${i}]`);
      checkRecord(entry, e, where);
      checkString(entry.styleName, `${e}.styleName`, where);
      checkHeadingLevel(entry.level, `${e}.level`, where);
    });
  }
  checkString(t.captionLabel, at("captionLabel"), where, optional);
  checkString(t.sequenceFieldIdentifier, at("sequenceFieldIdentifier"), where, optional);
  checkOneOf(t.leader, TAB_STOP_LEADERS, at("leader"), where, optional);
  checkOneOf(t.field, TOC_FIELD_MODES, at("field"), where, optional);
  checkNodeArray(t.cachedParagraphs, at("cachedParagraphs"), where, optional);
}

// -----------------------------------------------------------------------------
// Vocabulary at write
// -----------------------------------------------------------------------------

/** `ST_JcTable`: a table's alignment takes fewer values than a paragraph's. */
const TABLE_ALIGNMENTS: ReadonlySet<string> = new Set(["center", "end", "left", "right", "start"]);

export function checkTableAlignment(value: unknown, field: string, where: string): void {
  checkOneOf(value, TABLE_ALIGNMENTS, field, where, optional);
}

/** Widen a builder vocabulary by values the schema also allows, for writing what was read. */
function withValues(base: ReadonlySet<string>, extra: readonly string[]): ReadonlySet<string> {
  return new Set([...base, ...extra]);
}

/**
 * The vocabularies the writers enforce. Each is the schema's whole
 * enumeration — Transitional and Strict together, plus the values Word itself
 * writes beyond them (`w:hint="cs"`) — and `ooxml-vocabulary.test.ts` holds
 * each to the schema, so a value a conforming file holds is read into its
 * field. A value outside it is either kept as read (see {@link readToken}) or,
 * from a caller, refused rather than written as markup Word repairs. Where the
 * builder's type is narrower than the schema the writer set is wider: a cell
 * may be aligned `both` (`ST_VerticalJc`), and a Strict file's text
 * directions (`tb`, `rl`, …) are read as they are.
 */
export const WRITTEN = {
  underline: UNDERLINE_STYLES,
  highlight: HIGHLIGHT_COLORS,
  vertAlign: VERTICAL_ALIGNS,
  effect: TEXT_EFFECTS,
  emphasisMark: EMPHASIS_MARKS,
  combineBrackets: COMBINE_BRACKETS,
  fontHint: FONT_HINTS,
  shading: SHADING_TYPES,
  alignment: ALIGNMENTS,
  tableAlignment: TABLE_ALIGNMENTS,
  lineRule: LINE_SPACING_RULES,
  tabType: TAB_STOP_TYPES,
  tabLeader: TAB_STOP_LEADERS,
  textAlignment: TEXT_ALIGNMENTS,
  textDirection: withValues(TEXT_DIRECTIONS, ["tb", "rl", "lr", "tbV", "rlV", "lrV"]),
  dropCap: DROP_CAPS,
  frameWrap: FRAME_WRAPS,
  frameAnchor: FRAME_ANCHORS,
  frameXAlign: FRAME_X_ALIGNS,
  frameYAlign: FRAME_Y_ALIGNS,
  tableWidthType: TABLE_WIDTH_TYPES,
  tableLayout: TABLE_LAYOUTS,
  rowHeightRule: ROW_HEIGHT_RULES,
  cellVerticalAlign: withValues(VERTICAL_CELL_ALIGNS, ["both"]),
  sectionBreak: SECTION_BREAKS,
  pageVerticalAlign: PAGE_VERTICAL_ALIGNS,
  docGrid: DOC_GRID_TYPES,
  lineNumberRestart: LINE_NUMBER_RESTARTS,
  pageBorderDisplay: PAGE_BORDER_DISPLAYS,
  pageBorderOffset: PAGE_BORDER_OFFSETS,
  pageBorderZOrder: PAGE_BORDER_Z_ORDERS,
  horizontalRelative: HORIZONTAL_RELATIVE,
  verticalRelative: VERTICAL_RELATIVE,
  horizontalAlign: HORIZONTAL_ALIGNS,
  verticalAlign: VERTICAL_ALIGNS_ANCHOR,
  wrapSide: WRAP_SIDES,
  legendPosition: LEGEND_POSITIONS,
  dataLabelPosition: DATA_LABEL_POSITIONS
} as const;

/**
 * The attribute text a writer emits for `holder[key]`: the field, checked
 * against `vocabulary` — a {@link DocxArgumentError} naming `field`
 * when it is outside it — or, while the field is still as the reader left it,
 * the text the reader set aside for it (see {@link readToken}). `undefined`
 * when there is neither.
 */
export function writtenToken<T extends object>(
  holder: T,
  key: keyof T & string,
  vocabulary: ReadonlySet<string>,
  field: string
): string | undefined {
  const kept = unrecognizedValue(holder, key);
  if (kept !== undefined) {
    return kept;
  }
  const value = holder[key];
  if (value === undefined) {
    return undefined;
  }
  checkOneOf(value, vocabulary, field, MODEL_CHECK);
  return value as string;
}

/**
 * A value read from a file, returned when it is in `vocabulary`. Outside it the
 * field is left as an absent attribute leaves it — `undefined`, or `readAs`
 * for a field that cannot be absent — and the text is kept on `holder` under
 * `key`, so that writing the document puts it back as read instead of failing
 * on it or dropping it, and `Validation.validate` reports it. A read document
 * is not the caller's to fix, and a vocabulary narrower than the file's must
 * not lose what the file said.
 */
export function readToken(
  value: string | undefined,
  vocabulary: ReadonlySet<string>,
  holder: object,
  key: string,
  readAs?: string
): string | undefined {
  if (value === undefined) {
    return readAs;
  }
  if (vocabulary.has(value)) {
    return value;
  }
  setAsideValue(holder, key, value, { readAs });
  return readAs;
}
