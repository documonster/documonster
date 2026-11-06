/**
 * Argument checks for the PDF builder, editor and converters.
 *
 * `pdfNumber` writes a non-finite number as `0`, and a colour is read field by field, so
 * every wrong value used to become a *valid-looking* file that showed something other than
 * what was asked: a `NaN` font size set text at size 0, a colour string filled black, a
 * zero page width produced a page no viewer opens. Each check throws
 * {@link PdfArgumentError} naming the call and the field, before the call changes anything.
 *
 * An options object is checked against a schema (`@utils/schema`), pinned to its public
 * interface with `SchemaMatches`. The single-value checks below — and the colour and dash
 * checks — stay hand-written: the page builder runs them on every drawing call and
 * `Pdf.fromTable` on every cell, where a schema check measured several times slower.
 */

import { checkImageData } from "@pdf/builder/image-utils";
import { PdfArgumentError } from "@pdf/errors";
import type {
  PdfFontConfig,
  PdfFontFaces,
  PdfFontSource,
  PdfNamedFontFamily
} from "@pdf/font/font-config";
import type { PdfBook, PdfSheet } from "@pdf/pdf";
import { PageSizes } from "@pdf/types";
import type {
  PageSizeName,
  PdfColor,
  PdfDocumentOptions,
  PdfEncryptionOptions,
  PdfExportOptions,
  PdfHeaderFooterOptions,
  PdfImageWatermark,
  PdfPermissions,
  PdfTextWatermark,
  PdfWatermark,
  PdfWatermarkFilter
} from "@pdf/types";
import { isRecordArgument } from "@utils/object";
import { check, fail as schemaFail, s } from "@utils/schema";
import type { Flatten, Schema, SchemaCheckOptions, SchemaMatches } from "@utils/schema";
import { isLegalXmlChar } from "@utils/xml-encode";

function argumentError(field: string, message: string): PdfArgumentError {
  return new PdfArgumentError(field, message);
}

/** Where a checked value sits: the public call and the path to the value. */
export function at(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

function fail(where: string, field: string, expected: string, value: unknown): never {
  throw schemaFail(at(where, field), "", expected, value);
}

// ------------------------------------------------------------------ single values

/** A finite number. */
export function checkFinite(where: string, field: string, value: unknown): asserts value is number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail(where, field, "a finite number", value);
  }
}

/** A finite number, or `undefined`. */
export function checkOptionalFinite(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkFinite(where, field, value);
  }
}

/** A finite number `>= 0`, or `undefined`. */
export function checkOptionalNonNegative(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    fail(where, field, "a finite number >= 0", value);
  }
}

/** A finite number `> 0`, or `undefined`. */
export function checkOptionalPositive(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    fail(where, field, "a finite number > 0", value);
  }
}

/** An integer `>= min`, or `undefined`. */
export function checkOptionalInteger(where: string, field: string, value: unknown, min = 0): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "number" || !Number.isInteger(value) || value < min) {
    fail(where, field, `an integer >= ${min}`, value);
  }
}

/** A boolean, or `undefined`. */
export function checkOptionalBoolean(where: string, field: string, value: unknown): void {
  if (value !== undefined && typeof value !== "boolean") {
    fail(where, field, "a boolean", value);
  }
}

/** A string, or `undefined`. */
export function checkOptionalString(where: string, field: string, value: unknown): void {
  if (value !== undefined && typeof value !== "string") {
    fail(where, field, "a string", value);
  }
}

/** A string. */
export function checkString(where: string, field: string, value: unknown): asserts value is string {
  if (typeof value !== "string") {
    fail(where, field, "a string", value);
  }
}

/** One of `allowed`, or `undefined`. */
export function checkOptionalEnum(
  where: string,
  field: string,
  value: unknown,
  allowed: readonly string[]
): void {
  if (value !== undefined && (typeof value !== "string" || !allowed.includes(value))) {
    fail(where, field, `one of ${allowed.map(v => JSON.stringify(v)).join(", ")}`, value);
  }
}

/** A {@link PdfColor} — `r`, `g`, `b` and optional `a`, each from 0 to 1 — or `undefined`. */
export function checkOptionalColor(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, field, "a color { r, g, b, a? } with components from 0 to 1", value);
  }
  const color = value as Partial<Record<keyof PdfColor, unknown>>;
  for (const channel of ["r", "g", "b", "a"] as const) {
    const component = color[channel];
    if (
      (channel !== "a" || component !== undefined) &&
      (typeof component !== "number" || !(component >= 0 && component <= 1))
    ) {
      fail(where, `${field}.${channel}`, "a number from 0 to 1", component);
    }
  }
}

/** A rectangle `[x1, y1, x2, y2]` of finite numbers. */
export function checkRectTuple(where: string, field: string, value: unknown): void {
  if (!Array.isArray(value) || value.length !== 4) {
    fail(where, field, "an array [x1, y1, x2, y2]", value);
  }
  value.forEach((v, i) => checkFinite(where, `${field}[${i}]`, v));
}

/**
 * A dash pattern: finite lengths `>= 0`, not all zero (ISO 32000-1 §8.4.3.6 makes an
 * all-zero array an error), or `undefined`. An empty array means solid.
 */
export function checkOptionalDashPattern(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (!Array.isArray(value)) {
    fail(where, field, "an array of dash and gap lengths", value);
  }
  value.forEach((v, i) => checkOptionalNonNegative(where, `${field}[${i}]`, v));
  if (value.length > 0 && value.every(v => v === 0)) {
    fail(where, field, "lengths that are not all zero", "all zeros");
  }
}

/** An options object, or `undefined`. */
export function checkOptionalObject(where: string, field: string, value: unknown): void {
  if (
    value !== undefined &&
    (typeof value !== "object" || value === null || Array.isArray(value))
  ) {
    fail(where, field, "an options object", value);
  }
}

/** The written languages `textLanguage` accepts. */
export const CJK_LANGUAGES = ["zh-Hans", "zh-Hant", "ja", "ko"] as const;

/** A list of strings. */
export function checkStringList(where: string, field: string, value: unknown): void {
  check(stringListSchema, value, at(where, field));
}

/** Bytes, recognised structurally so a typed array from another realm is accepted. */
export function isByteArray(value: unknown): value is Uint8Array {
  return (
    ArrayBuffer.isView(value) && (value as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
  );
}

// ------------------------------------------------------------------ schema atoms

const finite = () => s.number({ finite: true });
const nonNegative = () => s.number({ finite: true, min: 0 });
const positive = () =>
  s
    .number({ finite: true, expected: "a finite number > 0" })
    .refine(v => v > 0, "a finite number > 0");
const unit = () => s.number({ min: 0, max: 1 });
const flag = () => s.boolean();
const optionsObject = { expected: "an options object" };
/** A key the checks do not read, named for its type. */
const unread = <T>() => s.unknown<T>().optional();

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** An object whose contents are someone else's to check. */
const anyObject = <T>(expected = "an options object") => s.custom<T>(isRecord, expected);

const validDate = () =>
  s.custom<Date>(v => v instanceof Date && !Number.isNaN(v.getTime()), "a valid Date");

const stringListSchema = s.readonlyArray(s.string(), { expected: "an array of strings" });

const colorSchema = s.object(
  { r: unit(), g: unit(), b: unit(), a: unit().optional() },
  { expected: "a color { r, g, b, a? } with components from 0 to 1" }
);
const _colorMatch: SchemaMatches<typeof colorSchema, PdfColor> = true;
void _colorMatch;

// ------------------------------------------------------------------ metadata

const metadataShape = {
  title: s.string().optional(),
  author: s.string().optional(),
  subject: s.string().optional(),
  keywords: s.string().optional(),
  creator: s.string().optional(),
  creationDate: validDate().optional(),
  modDate: validDate().optional()
};

/** The argument of `setMetadata`: the document properties, plus `producer`. */
const documentMetadataSchema = s.object(
  { ...metadataShape, producer: s.string().optional() },
  { expected: "an object of document properties" }
);

/**
 * The argument of `setMetadata` (builder and editor alike): an object whose string fields
 * are strings XML 1.0 can carry — the same text goes into an XMP packet, which cannot hold
 * a C0 control character, a lone surrogate or U+FFFE/U+FFFF — and whose dates are valid.
 */
export function checkDocumentMetadata(where: string, value: unknown): void {
  check(documentMetadataSchema, value, at(where, "metadata"));
  checkMetadataText(where, value, "metadata.");
}

/**
 * The metadata strings `Pdf.Builder#setMetadata` takes: text XML 1.0 can carry (see
 * {@link checkDocumentMetadata}). `prefix` is the path to the fields.
 */
export function checkMetadataText(
  where: string,
  metadata: Record<string, unknown>,
  prefix: string
): void {
  for (const field of ["title", "author", "subject", "keywords", "creator", "producer"]) {
    const text = metadata[field];
    if (typeof text !== "string") {
      continue;
    }
    for (const ch of text) {
      if (!isLegalXmlChar(ch.codePointAt(0)!)) {
        fail(
          where,
          `${prefix}${field}`,
          "text XML can carry (no control characters other than tab and newlines, no U+FFFE/U+FFFF, no lone surrogate)",
          `a string containing U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")}`
        );
      }
    }
  }
}

/**
 * The argument of `Editor#setMetadata`: what {@link checkDocumentMetadata} accepts, with
 * `null` also allowed for each field (it removes the property).
 */
export function checkMetadataUpdate(where: string, value: unknown): void {
  check(anyObject("an object of document properties"), value, at(where, "metadata"));
  const present: Record<string, unknown> = {};
  for (const [key, field] of Object.entries(value as Record<string, unknown>)) {
    if (field !== null) {
      present[key] = field;
    }
  }
  checkDocumentMetadata(where, present);
}

// ------------------------------------------------------------------ encryption

const encryptionSchema = s.object(
  {
    ownerPassword: s.string(),
    userPassword: s.string().optional(),
    permissions: anyObject<Partial<PdfPermissions>>().optional()
  },
  optionsObject
);
const _encryptionMatch: SchemaMatches<typeof encryptionSchema, PdfEncryptionOptions> = true;
void _encryptionMatch;

/**
 * Encryption options: `ownerPassword` a string, `userPassword` a string when given. `field` is
 * the path to the object (`encryption`, `options.encryption`).
 */
export function checkEncryptionOptions(where: string, field: string, value: unknown): void {
  check(encryptionSchema.optional(), value, at(where, field));
}

// ------------------------------------------------------------------ bytes

/** An `ArrayBuffer`, from this realm or another (where `instanceof` fails). */
function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === "[object ArrayBuffer]";
}

/**
 * Bytes a reader was given, as a `Uint8Array` over the same memory: a byte array as it is;
 * any other `ArrayBufferView` (a `DataView`, a Node `Buffer`, another typed array) and an
 * `ArrayBuffer` viewed, not copied — what `fetch().arrayBuffer()` and `FileReader` hand out,
 * and what `Workbook.read` and Word's `Document.read` already accept. `what` describes the bytes
 * in the error (`"the PDF file"`).
 */
export function readerBytes(
  where: string,
  field: string,
  value: unknown,
  what: string
): Uint8Array {
  if (ArrayBuffer.isView(value)) {
    if (Object.prototype.toString.call(value) === "[object Uint8Array]") {
      return value as Uint8Array;
    }
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (isArrayBuffer(value)) {
    return new Uint8Array(value);
  }
  return fail(
    where,
    field,
    `${what} as a Uint8Array, another ArrayBufferView or an ArrayBuffer`,
    value
  );
}

/**
 * A `Blob` (a browser `File` is one), recognised by its tag so one from another realm — an
 * iframe, a worker — is accepted as `instanceof Blob` would not.
 */
export function isBlob(value: unknown): value is Blob {
  const tag = Object.prototype.toString.call(value);
  return tag === "[object Blob]" || tag === "[object File]";
}

/**
 * The PDF file an async reader was given: bytes as {@link readerBytes} takes them, or a `Blob`,
 * whose bytes the caller reads.
 */
export function pdfBytesOrBlob(where: string, field: string, value: unknown): Uint8Array | Blob {
  if (isBlob(value)) {
    return value;
  }
  if (ArrayBuffer.isView(value) || isArrayBuffer(value)) {
    return readerBytes(where, field, value, "the PDF file");
  }
  return fail(
    where,
    field,
    "the PDF file as a Uint8Array, another ArrayBufferView, an ArrayBuffer or a Blob",
    value
  );
}

/** The bytes of a PDF file a reader was given — see {@link readerBytes}. */
export function pdfBytes(where: string, field: string, value: unknown): Uint8Array {
  return readerBytes(where, field, value, "the PDF file");
}

// ------------------------------------------------------------------ fonts

const FONT_BYTES = "the bytes of a .ttf or .ttc file";

/**
 * Font file bytes: a byte array (from any realm), not empty. A view of wider elements is
 * claimed here too, so it is refused as font bytes rather than as a `{ data }` record.
 */
const fontBytesSchema = s
  .bytes<Uint8Array>({ expected: FONT_BYTES })
  .refine(isByteArray, FONT_BYTES)
  .refine(v => v.byteLength > 0, "a non-empty .ttf or .ttc file");

const fontFaceRecordSchema = s.object(
  {
    data: fontBytesSchema,
    collectionIndex: s.number({ integer: true, min: 0 }).optional()
  },
  { readonly: true }
);

/** One `PdfFontSource`: font bytes, or `{ data, collectionIndex? }`. */
const fontSourceUnion = s.union([fontBytesSchema, fontFaceRecordSchema], {
  expected: "font bytes or { data, collectionIndex }"
});
const _fontSourceMatch: SchemaMatches<typeof fontSourceUnion, PdfFontSource> = true;
void _fontSourceMatch;
const fontSourceSchema = fontSourceUnion;

const fontFacesSchema = s.object(
  {
    regular: fontSourceSchema,
    bold: fontSourceSchema.optional(),
    italic: fontSourceSchema.optional(),
    boldItalic: fontSourceSchema.optional()
  },
  { expected: "an object of faces { regular, bold?, italic?, boldItalic? }", readonly: true }
);
const _fontFacesMatch: SchemaMatches<typeof fontFacesSchema, PdfFontFaces> = true;
void _fontFacesMatch;
const faces = fontFacesSchema;

/** A family name or alias: compared trimmed, so one of spaces names nothing. */
const familyName = () =>
  s.string().refine(name => fontFamilyKey(name) !== "", "a non-empty family name");

const fontFamilySchema = s.object(
  {
    name: familyName(),
    aliases: s.readonlyArray(familyName(), { expected: "an array of strings" }).optional(),
    faces
  },
  { expected: "a font family { name, aliases?, faces }", readonly: true }
);
const _fontFamilyMatch: SchemaMatches<typeof fontFamilySchema, PdfNamedFontFamily> = true;
void _fontFamilyMatch;

const fontConfigShapeSchema = s.object(
  {
    default: faces,
    families: s
      .readonlyArray(fontFamilySchema, {
        expected: "an array of font families"
      })
      .optional(),
    fallbackFamilies: s
      .readonlyArray(s.string(), { expected: "an array of family names" })
      .optional()
  },
  { expected: "a font configuration { default, families?, fallbackFamilies? }", readonly: true }
);
const _fontConfigMatch: SchemaMatches<typeof fontConfigShapeSchema, PdfFontConfig> = true;
void _fontConfigMatch;
const fontConfigSchema = fontConfigShapeSchema;

/** A family or alias name as `PdfFontConfig` takes it: compared trimmed and case-insensitively. */
function fontFamilyKey(name: string): string {
  return name.trim().replace(/\s+/g, " ").toLocaleLowerCase("en-US");
}

/**
 * The rules of a `PdfFontConfig` one entry cannot show alone: every family name and alias
 * unique, and every fallback family one of them. `config` has the right shape.
 */
function checkFontConfigNames(options: SchemaCheckOptions, config: PdfFontConfig): void {
  const claimed = new Set<string>();
  const claim = (path: string, name: string): void => {
    const key = fontFamilyKey(name);
    if (claimed.has(key)) {
      throw schemaFail(options, path, "a name no other family or alias uses", name);
    }
    claimed.add(key);
  };
  config.families?.forEach((family, i) => {
    claim(`.families[${i}].name`, family.name);
    family.aliases?.forEach((alias, j) => claim(`.families[${i}].aliases[${j}]`, alias));
  });
  config.fallbackFamilies?.forEach((name, i) => {
    if (!claimed.has(fontFamilyKey(name))) {
      throw schemaFail(
        options,
        `.fallbackFamilies[${i}]`,
        "the name or an alias of one of the configured families",
        name
      );
    }
  });
}

/**
 * A `PdfFontConfig` — the `fonts` option of every converter and of `Pdf.Builder`, and the
 * argument of `Pdf.Builder#embedFonts` — checked for shape before any face is parsed: an
 * object, `default` faces, named families with unique names and aliases, and
 * `fallbackFamilies` naming configured families. Bytes that are not a font are the parser's
 * to report (`PdfFontError`); everything a caller wrote wrong is a {@link PdfArgumentError}
 * naming the path, whichever entry point received it.
 */
export function checkPdfFontConfig(where: string, field: string, value: unknown): void {
  const options = at(where, field);
  check(fontConfigSchema, value, options);
  checkFontConfigNames(options, value);
}

/**
 * The face an `embedFont` call names, checked. `source` is font bytes or a
 * `{ data, collectionIndex }` record; the deprecated form passes the index as a second
 * argument, and its fields keep the names its signature gives them (`fontBytes`,
 * `collectionIndex`). `field` is the argument a later refusal of the face names.
 */
export function readFontSourceArgument(
  where: string,
  source: unknown,
  collectionIndex: unknown
): { data: Uint8Array; collectionIndex: number | undefined; field: string } {
  if (isRecordArgument(source)) {
    if (collectionIndex !== undefined) {
      throw new PdfArgumentError(
        "collectionIndex",
        `${where}: collectionIndex cannot be passed beside { data, collectionIndex } — give it as source.collectionIndex`
      );
    }
    check(fontSourceSchema, source, at(where, "source"));
    const face = source as { data: Uint8Array; collectionIndex?: number };
    return { data: face.data, collectionIndex: face.collectionIndex, field: "source" };
  }
  const field = collectionIndex === undefined ? "source" : "fontBytes";
  check(fontBytesSchema, source, at(where, field));
  checkOptionalInteger(where, "collectionIndex", collectionIndex, 0);
  return { data: source, collectionIndex: collectionIndex as number | undefined, field };
}

/**
 * The font options the builder, the editor and every converter share, with the same checks
 * wherever they are given. Pass them to `s.object` with the entry's other keys, then call
 * {@link checkFontOptionNames}.
 */
export const fontOptionsShape = {
  fonts: fontConfigSchema.optional(),
  preferSystemFonts: stringListSchema.optional(),
  textLanguage: s.oneOf(CJK_LANGUAGES).optional(),
  disableFontAutoDiscovery: flag().optional(),
  onWarning: s.function<(message: string) => void>().optional()
};

/** The cross-field rules of the font options, once their shape is checked. */
export function checkFontOptionNames(
  where: string,
  prefix: string,
  options: { readonly fonts?: PdfFontConfig } | undefined
): void {
  if (options?.fonts !== undefined) {
    checkFontConfigNames(at(where, `${prefix}fonts`), options.fonts);
  }
}

// ------------------------------------------------------------------ document options

const documentOptionsShape = {
  ...metadataShape,
  ...fontOptionsShape,
  encryption: encryptionSchema.optional()
};
const documentOptionsSchema = s.object(documentOptionsShape, optionsObject);
const _documentOptionsMatch: SchemaMatches<typeof documentOptionsSchema, PdfDocumentOptions> = true;
void _documentOptionsMatch;

/** The schema of a converter's options: {@link PdfDocumentOptions} and `shape`. */
export function documentOptionsSchemaWith<S extends Record<string, Schema<unknown, boolean>>>(
  shape: S
) {
  return s.object({ ...documentOptionsShape, ...shape }, optionsObject);
}

/**
 * {@link documentOptionsSchemaWith} for an options interface that declares its own keys
 * `readonly` while inheriting `PdfDocumentOptions`' mutable ones (`DocxToPdfOptions`).
 */
export function documentOptionsSchemaWithReadonly<
  S extends Record<string, Schema<unknown, boolean>>
>(shape: S) {
  return s.extend(documentOptionsSchema, shape, { readonly: true });
}

/**
 * Check a converter's options against `schema` (built with {@link documentOptionsSchemaWith}),
 * then the rules no single field shows. `options` is absent or an object of document options.
 */
export function checkConverterOptions(
  where: string,
  schema: Schema<unknown, boolean>,
  options: unknown
): void {
  check(schema.optional(), options, at(where, "options"));
  checkFontOptionNames(where, "options.", options as PdfDocumentOptions | undefined);
}

// ------------------------------------------------------------------ Pdf.fromTable input

const ARGB_HEX = /^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i;

function hexColor(example: string) {
  const expected = `an ARGB or RGB hex string such as "${example}"`;
  return s.string({ expected }).refine(v => ARGB_HEX.test(v), expected);
}

const CELL_ALIGNMENTS = ["left", "center", "right"];

/**
 * One cell value. Hand-written, because it runs once per cell of the table and a value's
 * kind decides the check (a `Date` is an object, but not a styled cell).
 */
function checkCellPrimitive(where: string, field: string, value: unknown): void {
  if (value === null || value === undefined || typeof value === "string") {
    return;
  }
  if (typeof value === "boolean") {
    return;
  }
  if (typeof value === "number") {
    // `NaN` and `Infinity` used to be printed as the words "NaN" and "Infinity".
    checkFinite(where, field, value);
    return;
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      fail(where, field, "a valid Date", value);
    }
    return;
  }
  // An object, function or bigint used to vanish from its row without a word.
  fail(where, field, "a string, finite number, boolean, valid Date, null or { value, … }", value);
}

function checkOptionalHexColor(where: string, field: string, value: unknown): void {
  if (value !== undefined && (typeof value !== "string" || !ARGB_HEX.test(value))) {
    // A colour name read as black.
    fail(where, field, 'an ARGB or RGB hex string such as "FFFF0000"', value);
  }
}

function checkPdfRows(where: string, field: string, rows: unknown): void {
  if (!Array.isArray(rows)) {
    fail(where, field, "an array of rows", rows);
  }
  rows.forEach((row: unknown, r) => {
    if (!Array.isArray(row)) {
      fail(where, `${field}[${r}]`, "an array of cell values", row);
    }
    row.forEach((cell: unknown, c) => {
      const cellAt = `${field}[${r}][${c}]`;
      if (typeof cell === "object" && cell !== null && !(cell instanceof Date)) {
        if (Array.isArray(cell) || !("value" in cell)) {
          fail(where, cellAt, "a cell value or a styled cell { value, … }", cell);
        }
        const styled = cell as Record<string, unknown>;
        checkCellPrimitive(where, `${cellAt}.value`, styled.value);
        checkOptionalBoolean(where, `${cellAt}.bold`, styled.bold);
        checkOptionalBoolean(where, `${cellAt}.italic`, styled.italic);
        checkOptionalPositive(where, `${cellAt}.fontSize`, styled.fontSize);
        checkOptionalHexColor(where, `${cellAt}.fontColor`, styled.fontColor);
        checkOptionalHexColor(where, `${cellAt}.fillColor`, styled.fillColor);
        checkOptionalEnum(where, `${cellAt}.align`, styled.align, CELL_ALIGNMENTS);
        return;
      }
      checkCellPrimitive(where, cellAt, cell);
    });
  });
}

const imageFormat = () => s.oneOf(["jpeg", "png"] as const);

const sheetImageSchema = s.object(
  {
    format: imageFormat(),
    data: s.custom<Uint8Array>(isByteArray, "a Uint8Array of image bytes"),
    col: s.number({ integer: true, min: 0 }),
    row: s.number({ integer: true, min: 0 }),
    width: positive(),
    height: positive()
  },
  optionsObject
);

const sheetSchema = s.object(
  {
    name: s.string().optional(),
    data: s.unknown<PdfSheet["data"]>(),
    columns: s
      .array(
        s.union(
          [
            positive(),
            s.object({ width: positive().optional(), header: s.string().optional() }, optionsObject)
          ],
          optionsObject
        ),
        { expected: "an array of column widths or { width, header }" }
      )
      .optional(),
    images: s.array(sheetImageSchema, { expected: "an array of images" }).optional()
  },
  optionsObject
);
const _sheetMatch: SchemaMatches<typeof sheetSchema, PdfSheet> = true;
void _sheetMatch;

const bookSchema = s.object({
  title: s.string().optional(),
  author: s.string().optional(),
  sheets: s.array(s.unknown<PdfSheet>(), { expected: "an array of sheets" })
});
const _bookMatch: SchemaMatches<typeof bookSchema, PdfBook> = true;
void _bookMatch;

function checkPdfSheet(where: string, field: string, value: unknown): void {
  check(sheetSchema, value, at(where, field));
  checkPdfRows(where, `${field}.data`, value.data);
  value.images?.forEach((image, i) =>
    checkImageData(where, `${field}.images[${i}].data`, image.data, image.format)
  );
}

/**
 * Check the input of `Pdf.fromTable`: rows, a sheet `{ data, … }` or a book `{ sheets }`.
 * Every cell is a string, finite number, boolean, valid `Date`, `null`/`undefined` or a
 * styled cell; colours are hex; sizes and widths are positive.
 */
export function checkPdfInput(where: string, input: unknown): void {
  if (Array.isArray(input)) {
    checkPdfRows(where, "input", input);
    return;
  }
  if (typeof input !== "object" || input === null) {
    fail(where, "input", "rows, a sheet { data } or a book { sheets }", input);
  }
  if ("sheets" in input) {
    check(bookSchema, input, at(where, "input"));
    input.sheets.forEach((sheet, i) => checkPdfSheet(where, `input.sheets[${i}]`, sheet));
    return;
  }
  checkPdfSheet(where, "input", input);
}

// ------------------------------------------------------------------ export options

const watermarkPlacementShape = {
  pages: s
    .readonlyArray(s.number({ integer: true, min: 1 }).orUndefined(), {
      expected: "an array of 1-based page numbers"
    })
    .optional(),
  sheets: stringListSchema.optional(),
  placement: s.oneOf(["under", "over"] as const).optional()
};

const watermarkCommonShape = {
  opacity: finite().optional(),
  rotation: finite().optional(),
  position: s
    .union([s.literal("center"), s.object({ x: finite(), y: finite() })], {
      expected: '"center" or { x, y }'
    })
    .optional(),
  repeat: unread<boolean>(),
  repeatSpacingX: finite().optional(),
  repeatSpacingY: finite().optional(),
  ...watermarkPlacementShape
};

const textWatermarkSchema = s.object(
  {
    type: s.literal("text"),
    text: s.string(),
    fontSize: finite().optional(),
    color: colorSchema.optional(),
    fontFamily: s.string().optional(),
    bold: unread<boolean>(),
    italic: unread<boolean>(),
    ...watermarkCommonShape
  },
  optionsObject
);
const _textWatermarkMatch: SchemaMatches<
  typeof textWatermarkSchema,
  Flatten<PdfTextWatermark & PdfWatermarkFilter>
> = true;
void _textWatermarkMatch;

const imageWatermarkSchema = s.object(
  {
    type: s.literal("image"),
    data: s.custom<Uint8Array>(isByteArray, "a Uint8Array of image bytes"),
    format: imageFormat(),
    scale: finite().optional(),
    width: finite().optional(),
    height: finite().optional(),
    ...watermarkCommonShape
  },
  optionsObject
);
const _imageWatermarkMatch: SchemaMatches<
  typeof imageWatermarkSchema,
  Flatten<PdfImageWatermark & PdfWatermarkFilter>
> = true;
void _imageWatermarkMatch;

const watermarkKindSchema = s.object({ type: s.oneOf(["text", "image"] as const) }, optionsObject);

function checkWatermark(where: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  const options = at(where, "options.watermark");
  check(watermarkKindSchema, value, options);
  if (value.type === "text") {
    check(textWatermarkSchema, value, options);
    return;
  }
  check(imageWatermarkSchema, value, options);
  checkImageData(where, "options.watermark.data", value.data, value.format);
}

const headerFooterSchema = s.object(
  {
    enabled: unread<boolean>(),
    fileName: s.string().optional(),
    filePath: s.string().optional(),
    date: validDate().optional(),
    locale: s.string().optional()
  },
  optionsObject
);
const _headerFooterMatch: SchemaMatches<typeof headerFooterSchema, PdfHeaderFooterOptions> = true;
void _headerFooterMatch;

const pageSizeNameExpected = `one of ${Object.keys(PageSizes)
  .map(v => JSON.stringify(v))
  .join(", ")}`;

/** A page dimension: a finite number first, then a positive one. */
const pageDimension = () => s.number({ finite: true }).refine(v => v > 0, "a finite number > 0");

const exportOptionsSchema = documentOptionsSchemaWith({
  pageSize: s
    .union(
      [
        s.member<PageSizeName>(Object.keys(PageSizes), { expected: pageSizeNameExpected }),
        s.object({ width: pageDimension(), height: pageDimension() }, optionsObject)
      ],
      optionsObject
    )
    .optional(),
  orientation: s.oneOf(["portrait", "landscape"] as const).optional(),
  margins: s
    .object(
      {
        top: nonNegative().optional(),
        right: nonNegative().optional(),
        bottom: nonNegative().optional(),
        left: nonNegative().optional()
      },
      optionsObject
    )
    .optional(),
  sheets: s
    .readonlyArray(
      // Positions are 1-based: `0` used to pass this check and then match no sheet.
      s.custom<string | number>(
        v => typeof v === "string" || (typeof v === "number" && Number.isInteger(v) && v >= 1),
        "a sheet name or a 1-based sheet position (an integer >= 1)"
      ),
      { expected: "an array of sheet names or 1-based positions" }
    )
    .optional(),
  ignorePrintArea: unread<boolean>(),
  fitToPage: unread<boolean>(),
  scale: finite().optional(),
  fitToWidth: nonNegative().optional(),
  fitToHeight: nonNegative().optional(),
  repeatCols: repeatCount(),
  showRowColHeaders: unread<boolean>(),
  pageOrder: s.oneOf(["downThenOver", "overThenDown"] as const).optional(),
  blackAndWhite: unread<boolean>(),
  draft: unread<boolean>(),
  errors: s.oneOf(["displayed", "blank", "dash", "NA"] as const).optional(),
  cellComments: s.oneOf(["none", "atEnd", "asDisplayed"] as const).optional(),
  showGridLines: unread<boolean>(),
  gridLineColor: hexColor("FFD0D0D0").optional(),
  horizontalCentered: unread<boolean>(),
  verticalCentered: unread<boolean>(),
  repeatRows: repeatCount(),
  defaultFontFamily: s.string().optional(),
  defaultFontSize: positive().optional(),
  showSheetNames: unread<boolean>(),
  showPageNumbers: unread<boolean>(),
  headerFooter: headerFooterSchema.optional(),
  font: s.custom<Uint8Array>(isByteArray, "a Uint8Array").optional(),
  watermark: unread<PdfWatermark>()
});
const _exportOptionsMatch: SchemaMatches<typeof exportOptionsSchema, PdfExportOptions> = true;
void _exportOptionsMatch;

/** `repeatRows` / `repeatCols`: a count, or `false` for none. */
function repeatCount() {
  return s
    .custom<number | false>(
      v => v === false || (typeof v === "number" && Number.isInteger(v) && v >= 0),
      "an integer >= 0"
    )
    .optional();
}

/**
 * Check every {@link PdfExportOptions} field `Pdf.fromTable` and `Pdf.fromExcel` read.
 *
 * A value outside a documented range that the exporter has always clamped (`scale`, a
 * watermark's `opacity` or `fontSize`) keeps being clamped; what is refused is a value that
 * has no reading at all — `NaN`, a string where a number belongs, a name outside its
 * enumeration — which used to reach the file as `0`, as a default, or as `NaN` layout.
 */
export function checkPdfExportOptions(where: string, options: PdfExportOptions | undefined): void {
  checkConverterOptions(where, exportOptionsSchema, options);
  if (options === undefined) {
    return;
  }
  if (options.font !== undefined && options.fonts !== undefined) {
    throw schemaFail(
      at(where, "options"),
      ".fonts",
      "omitted when options.font is given — options.fonts names every face",
      options.fonts
    );
  }
  checkWatermark(where, options.watermark);
}

/**
 * A page a draw surface paints: a page builder (`Pdf.Builder#addPage`, `Pdf.Editor#getPage`)
 * or anything with the six marks it puts down.
 */
export function checkDrawPage(where: string, page: unknown): void {
  for (const method of [
    "drawRect",
    "drawEllipse",
    "drawPath",
    "drawLine",
    "drawText",
    "getContentStream"
  ] as const) {
    if (typeof (page as Record<string, unknown> | null)?.[method] !== "function") {
      throw new PdfArgumentError(
        "page",
        `${where}: page must be a page builder (Pdf.Builder#addPage, Pdf.Editor#getPage) — it has no ${method}()`
      );
    }
  }
}
