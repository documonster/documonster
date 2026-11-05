/**
 * Argument checks for the PDF builder, editor and converters.
 *
 * `pdfNumber` writes a non-finite number as `0`, and a colour is read field by field, so
 * every wrong value used to become a *valid-looking* file that showed something other than
 * what was asked: a `NaN` font size set text at size 0, a colour string filled black, a
 * zero page width produced a page no viewer opens. Each check throws
 * {@link PdfArgumentError} naming the call and the field, before the call changes anything.
 */

import { checkImageData } from "@pdf/builder/image-utils";
import { PdfArgumentError } from "@pdf/errors";
import { PageSizes } from "@pdf/types";
import type { PdfColor, PdfDocumentOptions, PdfExportOptions } from "@pdf/types";
import { describeValue as describe } from "@utils/describe-value";
import { isRecordArgument } from "@utils/object";
import { isLegalXmlChar } from "@utils/xml-encode";

function fail(where: string, field: string, expected: string, value: unknown): never {
  throw new PdfArgumentError(
    field,
    `${where}: ${field} must be ${expected}, got ${describe(value)}`
  );
}

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

/** A number in `[0, 1]`, or `undefined`. */
export function checkOptionalUnit(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "number" || !(value >= 0 && value <= 1)) {
    fail(where, field, "a number from 0 to 1", value);
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
  for (const channel of ["r", "g", "b"] as const) {
    if (typeof color[channel] !== "number" || !(color[channel] >= 0 && color[channel] <= 1)) {
      fail(where, `${field}.${channel}`, "a number from 0 to 1", color[channel]);
    }
  }
  checkOptionalUnit(where, `${field}.a`, color.a);
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

/** A valid `Date`, or `undefined`. */
export function checkOptionalDate(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    fail(where, field, "a valid Date", value instanceof Date ? "an invalid Date" : value);
  }
}

/** A page size: `width` and `height` finite and `> 0`; `prefix` is the path to them (`options.`). */
export function checkPageDimensions(
  where: string,
  width: unknown,
  height: unknown,
  prefix: string
): void {
  checkOptionalPositive(where, `${prefix}width`, width);
  checkOptionalPositive(where, `${prefix}height`, height);
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
export const CJK_LANGUAGES: readonly string[] = ["zh-Hans", "zh-Hant", "ja", "ko"];

/** A list of strings. */
export function checkStringList(where: string, field: string, value: unknown): void {
  if (!Array.isArray(value)) {
    fail(where, field, "an array of strings", value);
  }
  value.forEach((v, i) => checkString(where, `${field}[${i}]`, v));
}

/** Page options: `width` and `height`, each finite and `> 0` when given. */
export function checkPageOptions(
  where: string,
  options: { width?: unknown; height?: unknown } | undefined
): void {
  checkOptionalObject(where, "options", options);
  checkPageDimensions(where, options?.width, options?.height, "options.");
}

/**
 * Document metadata — `title`, `author`, `subject`, `keywords`, `creator` strings and valid
 * dates. `prefix` is the path to the fields (`metadata.`, `options.`).
 */
export function checkMetadata(
  where: string,
  metadata: Record<string, unknown> | undefined,
  prefix: string
): void {
  if (metadata === undefined) {
    return;
  }
  for (const field of ["title", "author", "subject", "keywords", "creator"]) {
    checkOptionalString(where, `${prefix}${field}`, metadata[field]);
  }
  checkOptionalDate(where, `${prefix}creationDate`, metadata.creationDate);
  checkOptionalDate(where, `${prefix}modDate`, metadata.modDate);
}

/**
 * The argument of `setMetadata` (builder and editor alike): an object whose string fields
 * are strings XML 1.0 can carry — the same text goes into an XMP packet, which cannot hold
 * a C0 control character, a lone surrogate or U+FFFE/U+FFFF — and whose dates are valid.
 */
export function checkDocumentMetadata(where: string, value: unknown): void {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, "metadata", "an object of document properties", value);
  }
  const metadata = value as Record<string, unknown>;
  checkMetadata(where, metadata, "metadata.");
  checkOptionalString(where, "metadata.producer", metadata.producer);
  checkMetadataText(where, metadata, "metadata.");
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
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, "metadata", "an object of document properties", value);
  }
  const present: Record<string, unknown> = {};
  for (const [key, field] of Object.entries(value as Record<string, unknown>)) {
    if (field !== null) {
      present[key] = field;
    }
  }
  checkDocumentMetadata(where, present);
}

/**
 * Encryption options: `ownerPassword` a string, `userPassword` a string when given. `field` is
 * the path to the object (`encryption`, `options.encryption`).
 */
export function checkEncryptionOptions(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  checkOptionalObject(where, field, value);
  const encryption = value as Record<string, unknown>;
  checkString(where, `${field}.ownerPassword`, encryption.ownerPassword);
  checkOptionalString(where, `${field}.userPassword`, encryption.userPassword);
  checkOptionalObject(where, `${field}.permissions`, encryption.permissions);
}

/** Bytes, recognised structurally so a typed array from another realm is accepted. */
export function isByteArray(value: unknown): value is Uint8Array {
  return (
    ArrayBuffer.isView(value) && (value as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
  );
}

function checkOptionalBytes(where: string, field: string, value: unknown): void {
  if (value !== undefined && !isByteArray(value)) {
    fail(where, field, "a Uint8Array", value);
  }
}

/** An `ArrayBuffer`, from this realm or another (where `instanceof` fails). */
function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === "[object ArrayBuffer]";
}

/**
 * Bytes a reader was given, as a `Uint8Array` over the same memory: a byte array as it is;
 * any other `ArrayBufferView` (a `DataView`, a Node `Buffer`, another typed array) and an
 * `ArrayBuffer` viewed, not copied — what `fetch().arrayBuffer()` and `FileReader` hand out,
 * and what `Workbook.read` and Word's `Io.read` already accept. `what` describes the bytes
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

/** Font file bytes: a `Uint8Array` (or another realm's), not empty. */
export function checkFontBytes(where: string, value: unknown, field = "fontBytes"): void {
  if (!isByteArray(value) || value.byteLength === 0) {
    fail(where, field, "the bytes of a .ttf or .ttc file", isByteArray(value) ? "no bytes" : value);
  }
}

/** One `PdfFontSource`: font bytes, or `{ data, collectionIndex? }`. */
export function checkFontSource(where: string, field: string, value: unknown): void {
  if (isByteArray(value)) {
    checkFontBytes(where, value, field);
    return;
  }
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, field, "font bytes or { data, collectionIndex }", value);
  }
  const source = value as Record<string, unknown>;
  checkFontBytes(where, source.data, `${field}.data`);
  checkOptionalInteger(where, `${field}.collectionIndex`, source.collectionIndex, 0);
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
    checkFontSource(where, "source", source);
    const face = source as { data: Uint8Array; collectionIndex?: number };
    return { data: face.data, collectionIndex: face.collectionIndex, field: "source" };
  }
  const field = collectionIndex === undefined ? "source" : "fontBytes";
  checkFontBytes(where, source, field);
  checkOptionalInteger(where, "collectionIndex", collectionIndex, 0);
  return {
    data: source as Uint8Array,
    collectionIndex: collectionIndex as number | undefined,
    field
  };
}

/** A `PdfFontFaces`: `regular` and the optional styled faces. */
function checkFontFaces(where: string, field: string, value: unknown): void {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, field, "an object of faces { regular, bold?, italic?, boldItalic? }", value);
  }
  const faces = value as Record<string, unknown>;
  checkFontSource(where, `${field}.regular`, faces.regular);
  for (const style of ["bold", "italic", "boldItalic"]) {
    if (faces[style] !== undefined) {
      checkFontSource(where, `${field}.${style}`, faces[style]);
    }
  }
}

/** A family or alias name as `PdfFontConfig` takes it: compared trimmed and case-insensitively. */
function fontFamilyKey(name: string): string {
  return name.trim().replace(/\s+/g, " ").toLocaleLowerCase("en-US");
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
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, field, "a font configuration { default, families?, fallbackFamilies? }", value);
  }
  const config = value as Record<string, unknown>;
  checkFontFaces(where, `${field}.default`, config.default);
  const claimed = new Set<string>();
  const claim = (at: string, name: unknown): void => {
    checkString(where, at, name);
    const key = fontFamilyKey(name);
    if (key === "") {
      fail(where, at, "a non-empty family name", name);
    }
    if (claimed.has(key)) {
      fail(where, at, "a name no other family or alias uses", name);
    }
    claimed.add(key);
  };
  if (config.families !== undefined) {
    if (!Array.isArray(config.families)) {
      fail(where, `${field}.families`, "an array of font families", config.families);
    }
    config.families.forEach((family: unknown, i) => {
      const at = `${field}.families[${i}]`;
      if (typeof family !== "object" || family === null || Array.isArray(family)) {
        fail(where, at, "a font family { name, aliases?, faces }", family);
      }
      const f = family as Record<string, unknown>;
      claim(`${at}.name`, f.name);
      if (f.aliases !== undefined) {
        if (!Array.isArray(f.aliases)) {
          fail(where, `${at}.aliases`, "an array of strings", f.aliases);
        }
        f.aliases.forEach((alias: unknown, j) => claim(`${at}.aliases[${j}]`, alias));
      }
      checkFontFaces(where, `${at}.faces`, f.faces);
    });
  }
  if (config.fallbackFamilies !== undefined) {
    if (!Array.isArray(config.fallbackFamilies)) {
      fail(where, `${field}.fallbackFamilies`, "an array of family names", config.fallbackFamilies);
    }
    config.fallbackFamilies.forEach((name: unknown, i) => {
      const at = `${field}.fallbackFamilies[${i}]`;
      checkString(where, at, name);
      if (!claimed.has(fontFamilyKey(name))) {
        fail(where, at, "the name or an alias of one of the configured families", name);
      }
    });
  }
}

export function checkOptionalFunction(where: string, field: string, value: unknown): void {
  if (value !== undefined && typeof value !== "function") {
    fail(where, field, "a function", value);
  }
}

const ARGB_HEX = /^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i;

const CELL_ALIGNMENTS = ["left", "center", "right"];

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
    checkOptionalDate(where, field, value);
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
      const at = `${field}[${r}][${c}]`;
      if (typeof cell === "object" && cell !== null && !(cell instanceof Date)) {
        if (Array.isArray(cell) || !("value" in cell)) {
          fail(where, at, "a cell value or a styled cell { value, … }", cell);
        }
        const styled = cell as Record<string, unknown>;
        checkCellPrimitive(where, `${at}.value`, styled.value);
        if (styled.bold !== undefined && typeof styled.bold !== "boolean") {
          fail(where, `${at}.bold`, "a boolean", styled.bold);
        }
        if (styled.italic !== undefined && typeof styled.italic !== "boolean") {
          fail(where, `${at}.italic`, "a boolean", styled.italic);
        }
        checkOptionalPositive(where, `${at}.fontSize`, styled.fontSize);
        checkOptionalHexColor(where, `${at}.fontColor`, styled.fontColor);
        checkOptionalHexColor(where, `${at}.fillColor`, styled.fillColor);
        checkOptionalEnum(where, `${at}.align`, styled.align, CELL_ALIGNMENTS);
        return;
      }
      checkCellPrimitive(where, at, cell);
    });
  });
}

function checkPdfSheet(where: string, field: string, value: unknown): void {
  checkOptionalObject(where, field, value ?? null);
  const sheet = value as Record<string, unknown>;
  checkOptionalString(where, `${field}.name`, sheet.name);
  checkPdfRows(where, `${field}.data`, sheet.data);
  if (sheet.columns !== undefined) {
    if (!Array.isArray(sheet.columns)) {
      fail(
        where,
        `${field}.columns`,
        "an array of column widths or { width, header }",
        sheet.columns
      );
    }
    sheet.columns.forEach((column: unknown, i) => {
      const at = `${field}.columns[${i}]`;
      if (typeof column === "number") {
        // A width of 0 or less laid every cell of the column out with no room, and the
        // text vanished.
        checkOptionalPositive(where, at, column);
        return;
      }
      checkOptionalObject(where, at, column ?? null);
      const c = column as Record<string, unknown>;
      checkOptionalPositive(where, `${at}.width`, c.width);
      checkOptionalString(where, `${at}.header`, c.header);
    });
  }
  if (sheet.images !== undefined) {
    if (!Array.isArray(sheet.images)) {
      fail(where, `${field}.images`, "an array of images", sheet.images);
    }
    sheet.images.forEach((image: unknown, i) => {
      const at = `${field}.images[${i}]`;
      checkOptionalObject(where, at, image ?? null);
      const img = image as Record<string, unknown>;
      checkOptionalEnum(where, `${at}.format`, img.format, ["jpeg", "png"]);
      if (img.format === undefined) {
        fail(where, `${at}.format`, '"jpeg" or "png"', undefined);
      }
      if (!isByteArray(img.data)) {
        fail(where, `${at}.data`, "a Uint8Array of image bytes", img.data);
      }
      checkImageData(where, `${at}.data`, img.data, img.format as "jpeg" | "png");
      checkOptionalInteger(where, `${at}.col`, img.col, 0);
      checkOptionalInteger(where, `${at}.row`, img.row, 0);
      if (img.col === undefined || img.row === undefined) {
        fail(where, `${at}.${img.col === undefined ? "col" : "row"}`, "an integer >= 0", undefined);
      }
      checkOptionalPositive(where, `${at}.width`, img.width);
      checkOptionalPositive(where, `${at}.height`, img.height);
      if (img.width === undefined || img.height === undefined) {
        fail(
          where,
          `${at}.${img.width === undefined ? "width" : "height"}`,
          "a finite number > 0",
          undefined
        );
      }
    });
  }
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
    const book = input as Record<string, unknown>;
    checkOptionalString(where, "input.title", book.title);
    checkOptionalString(where, "input.author", book.author);
    if (!Array.isArray(book.sheets)) {
      fail(where, "input.sheets", "an array of sheets", book.sheets);
    }
    book.sheets.forEach((sheet: unknown, i) => checkPdfSheet(where, `input.sheets[${i}]`, sheet));
    return;
  }
  checkPdfSheet(where, "input", input);
}

/**
 * The font options the builder, the editor and every converter share — `fonts`,
 * `preferSystemFonts`, `textLanguage`, `disableFontAutoDiscovery`, `onWarning` — with the
 * same checks wherever they are given. `prefix` is the path to them (`options.`).
 */
export function checkFontOptions(
  where: string,
  prefix: string,
  options:
    | {
        readonly fonts?: unknown;
        readonly preferSystemFonts?: unknown;
        readonly textLanguage?: unknown;
        readonly disableFontAutoDiscovery?: unknown;
        readonly onWarning?: unknown;
      }
    | undefined
): void {
  if (options?.fonts !== undefined) {
    checkPdfFontConfig(where, `${prefix}fonts`, options.fonts);
  }
  if (options?.preferSystemFonts !== undefined) {
    checkStringList(where, `${prefix}preferSystemFonts`, options.preferSystemFonts);
  }
  checkOptionalEnum(where, `${prefix}textLanguage`, options?.textLanguage, CJK_LANGUAGES);
  checkOptionalBoolean(
    where,
    `${prefix}disableFontAutoDiscovery`,
    options?.disableFontAutoDiscovery
  );
  checkOptionalFunction(where, `${prefix}onWarning`, options?.onWarning);
}

/**
 * Check the {@link PdfDocumentOptions} every converter takes — metadata, encryption and the
 * font settings — the same way for each. `options` was checked to be an object (or absent).
 */
export function checkPdfDocumentOptions(
  where: string,
  options: PdfDocumentOptions | undefined
): void {
  if (options === undefined) {
    return;
  }
  checkMetadata(where, options as Record<string, unknown>, "options.");
  checkFontOptions(where, "options.", options);
  checkEncryptionOptions(where, "options.encryption", options.encryption);
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
  checkOptionalObject(where, "options", options);
  if (options === undefined) {
    return;
  }
  const o = options as Record<keyof PdfExportOptions, unknown>;
  if (typeof o.pageSize === "string") {
    checkOptionalEnum(where, "options.pageSize", o.pageSize, Object.keys(PageSizes));
  } else if (o.pageSize !== undefined) {
    checkOptionalObject(where, "options.pageSize", o.pageSize);
    const size = o.pageSize as Record<string, unknown>;
    checkFinite(where, "options.pageSize.width", size.width);
    checkFinite(where, "options.pageSize.height", size.height);
    checkPageDimensions(where, size.width, size.height, "options.pageSize.");
  }
  checkOptionalEnum(where, "options.orientation", o.orientation, ["portrait", "landscape"]);
  checkOptionalObject(where, "options.margins", o.margins);
  for (const side of ["top", "right", "bottom", "left"]) {
    checkOptionalNonNegative(
      where,
      `options.margins.${side}`,
      (o.margins as Record<string, unknown>)?.[side]
    );
  }
  if (o.sheets !== undefined) {
    if (!Array.isArray(o.sheets)) {
      fail(where, "options.sheets", "an array of sheet names or 1-based positions", o.sheets);
    }
    o.sheets.forEach((sheet: unknown, i) => {
      // Positions are 1-based: `0` used to pass this check and then match no sheet.
      if (
        typeof sheet !== "string" &&
        (typeof sheet !== "number" || !Number.isInteger(sheet) || sheet < 1)
      ) {
        fail(
          where,
          `options.sheets[${i}]`,
          "a sheet name or a 1-based sheet position (an integer >= 1)",
          sheet
        );
      }
    });
  }
  checkOptionalFinite(where, "options.scale", o.scale);
  checkOptionalNonNegative(where, "options.fitToWidth", o.fitToWidth);
  checkOptionalNonNegative(where, "options.fitToHeight", o.fitToHeight);
  for (const field of ["repeatRows", "repeatCols"] as const) {
    if (o[field] !== false) {
      checkOptionalInteger(where, `options.${field}`, o[field], 0);
    }
  }
  checkOptionalEnum(where, "options.pageOrder", o.pageOrder, ["downThenOver", "overThenDown"]);
  checkOptionalEnum(where, "options.errors", o.errors, ["displayed", "blank", "dash", "NA"]);
  checkOptionalEnum(where, "options.cellComments", o.cellComments, [
    "none",
    "atEnd",
    "asDisplayed"
  ]);
  if (o.gridLineColor !== undefined) {
    if (typeof o.gridLineColor !== "string" || !ARGB_HEX.test(o.gridLineColor)) {
      fail(
        where,
        "options.gridLineColor",
        'an ARGB or RGB hex string such as "FFD0D0D0"',
        o.gridLineColor
      );
    }
  }
  checkOptionalString(where, "options.defaultFontFamily", o.defaultFontFamily);
  checkOptionalPositive(where, "options.defaultFontSize", o.defaultFontSize);
  checkOptionalObject(where, "options.headerFooter", o.headerFooter);
  const headerFooter = o.headerFooter as Record<string, unknown> | undefined;
  for (const field of ["fileName", "filePath", "locale"]) {
    checkOptionalString(where, `options.headerFooter.${field}`, headerFooter?.[field]);
  }
  checkOptionalDate(where, "options.headerFooter.date", headerFooter?.date);
  checkPdfDocumentOptions(where, options);
  checkOptionalBytes(where, "options.font", o.font);
  if (o.font !== undefined && o.fonts !== undefined) {
    throw new PdfArgumentError(
      "options.fonts",
      `${where}: options.font and options.fonts cannot be used together — options.fonts names every face`
    );
  }
  checkWatermark(where, o.watermark);
}

function checkWatermark(where: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  checkOptionalObject(where, "options.watermark", value);
  const wm = value as Record<string, unknown>;
  checkOptionalEnum(where, "options.watermark.type", wm.type, ["text", "image"]);
  if (wm.type === undefined) {
    fail(where, "options.watermark.type", '"text" or "image"', undefined);
  }
  if (wm.type === "text") {
    checkString(where, "options.watermark.text", wm.text);
    checkOptionalFinite(where, "options.watermark.fontSize", wm.fontSize);
    checkOptionalColor(where, "options.watermark.color", wm.color);
    checkOptionalString(where, "options.watermark.fontFamily", wm.fontFamily);
  } else {
    if (!isByteArray(wm.data)) {
      fail(where, "options.watermark.data", "a Uint8Array of image bytes", wm.data);
    }
    checkOptionalEnum(where, "options.watermark.format", wm.format, ["jpeg", "png"]);
    if (wm.format === undefined) {
      fail(where, "options.watermark.format", '"jpeg" or "png"', undefined);
    }
    checkImageData(where, "options.watermark.data", wm.data, wm.format as "jpeg" | "png");
    checkOptionalFinite(where, "options.watermark.scale", wm.scale);
    checkOptionalFinite(where, "options.watermark.width", wm.width);
    checkOptionalFinite(where, "options.watermark.height", wm.height);
  }
  checkOptionalFinite(where, "options.watermark.opacity", wm.opacity);
  checkOptionalFinite(where, "options.watermark.rotation", wm.rotation);
  checkOptionalFinite(where, "options.watermark.repeatSpacingX", wm.repeatSpacingX);
  checkOptionalFinite(where, "options.watermark.repeatSpacingY", wm.repeatSpacingY);
  if (wm.position !== undefined && wm.position !== "center") {
    if (typeof wm.position !== "object" || wm.position === null) {
      fail(where, "options.watermark.position", '"center" or { x, y }', wm.position);
    }
    const position = wm.position as Record<string, unknown>;
    checkFinite(where, "options.watermark.position.x", position.x);
    checkFinite(where, "options.watermark.position.y", position.y);
  }
  if (wm.pages !== undefined) {
    if (!Array.isArray(wm.pages)) {
      fail(where, "options.watermark.pages", "an array of 1-based page numbers", wm.pages);
    }
    wm.pages.forEach((page: unknown, i) =>
      checkOptionalInteger(where, `options.watermark.pages[${i}]`, page, 1)
    );
  }
  if (wm.sheets !== undefined) {
    checkStringList(where, "options.watermark.sheets", wm.sheets);
  }
  checkOptionalEnum(where, "options.watermark.placement", wm.placement, ["under", "over"]);
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
