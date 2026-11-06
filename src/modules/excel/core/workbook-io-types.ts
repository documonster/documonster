/**
 * The public IO option and report types, in one place for both platforms.
 *
 * **Why they are not declared where they are used.** `xlsx-io.ts` and `xlsx-io.browser.ts` are platform
 * variants, and these three types are not: a `format` option means the same thing in a browser as it does
 * under Node. Declaring them twice made them two copies of one public contract, and they had already
 * drifted — the write options' documentation differed between the two files, so a consumer reading the
 * browser declaration was told something the Node one did not say. Adding a field to only one of them is
 * the failure this file exists to make impossible.
 *
 * What genuinely differs stays in the variants: `toBuffer` returns a `Buffer` on Node and a
 * `Uint8Array` in a browser, the stream types are each platform's own, and file-path IO is Node-only.
 */
import type { EncryptOptions } from "@archive/office-crypto/agile";
import type { UnrecognizedValue } from "@excel/core/unrecognized-values";
import { keptSettingsReport } from "@excel/core/unrecognized-values";
import type { WorkbookData } from "@excel/core/workbook-core";
import { createOptionsSchema } from "@excel/core/workbook-create-options";
import type { WorkbookCreateOptions } from "@excel/core/workbook-create-options";
import type { WorkbookFormat } from "@excel/core/workbook-format";
import { isArrayBufferLike, xlsbReadOptions } from "@excel/core/workbook-format";
import type { XlsxStreamOptions } from "@excel/core/xlsx-stream";
import { ExcelArgumentError } from "@excel/errors";
import type {} from "@excel/utils/arguments";
import { argumentChecks, argumentFailure, checkSchema } from "@excel/utils/arguments";
import { decryptionOf, openWorkbookPackage } from "@excel/utils/container-errors";
import type { WorkbookDecryption } from "@excel/utils/container-errors";
import type {
  XlsxParseStream,
  XlsxReadOptions,
  XlsxWriteOptions,
  XlsxZipReadLimits,
  ZipWriterOptions
} from "@excel/xlsx/types";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
export type { WorkbookCreateOptions } from "@excel/core/workbook-create-options";

/**
 * Options common to reading either format.
 *
 * `format` is declared rather than inferred so a caller can be explicit, and so a misspelling is a
 * compile error — these option bags carry no index signature for exactly that reason.
 */
export interface WorkbookReadOptions extends XlsxReadOptions {
  /** Force a format instead of detecting one. */
  readonly format?: WorkbookFormat;
  /**
   * What to do about content the XLSB reader could not recover.
   *
   * `"ignore"` (the default) reads the workbook without it. `"error"` refuses and names what was lost, on
   * the thrown error's `items`. Ignored when reading XLSX.
   *
   * The default is the opposite of the write side's on purpose: writing a workbook this library holds in
   * memory can afford to stop, while refusing to *read* a real file because part of it uses a record
   * whose layout is unestablished would make the reader unusable. See `readXlsbInto`.
   */
  readonly unsupported?: "error" | "ignore";
  /**
   * What to do with cells that carry formatting and no value.
   *
   * - `"keep"` (the default) gives each one a cell in the model, which is what a caller iterating cells expects.
   * - `"collapse"` accumulates them into rectangles instead. Excel writes a `BrtCellBlank` for every cell of a
   *   formatted region, so a sheet with a formatted column has one per row to the sheet's end — measured at 186 MB of
   *   retained heap and 16,379 physical rows against 253 rows of actual data, against 0.4 MB and 253 rows collapsed.
   *
   * **`"collapse"` is lossless in both containers.** The rectangles describe exactly the cells they came from — a
   * `BrtCellBlank` run in XLSB, a `<c s="N"/>` run in XLSX — so writing the workbook back in the container it was read
   * from reproduces them byte for byte and nothing is owed to a fidelity report. Measured on XLSX at 43.8 MB and 8,000
   * physical rows against 200 rows of data, writing back to an identical 148,820 bytes.
   *
   * **Crossing containers costs nothing either, and the warning that used to sit here was wrong.** It said "a collapsed
   * read written as XLSX loses the blank formatting. That loss is currently silent". Measured on 300 styled blanks in one
   * column, all four combinations — `keep`/`collapse` read, written as XLSX/XLSB — come back with 300 formatted blanks and
   * the same fill on each. Both writers expand the rectangles; the XLSX one does it in `worksheet-xform`, the XLSB one in
   * `model-adapter`. There is no report owed because there is no loss.
   */
  readonly blankCells?: "keep" | "collapse";
  /**
   * What to do with a formula cell's expression when reading XLSB.
   *
   * A formula cell in XLSB carries two independent things: the value Excel last computed, and the token stream that
   * computed it. The value always decodes. The token stream is the hard part of the format, and the part where this
   * codec can be defeated by a construct it has not met — so what should happen then is a policy, not a default.
   *
   * - `"preserve"` (the default) keeps every expression it can decode and the cached value where it cannot, listing the
   *   address under `undecodedFormulas`. Nothing throws and nothing is invented, which is what a caller opening an
   *   arbitrary file wants.
   * - `"cached"` does not decode expressions at all: every formula cell becomes its cached value. Choose it to pull
   *   numbers out of a large workbook without paying for token decoding, or when Excel's computed value is worth more
   *   to you than an expression reconstructed by this library.
   * - `"error"` throws `XlsbFormulaDecodeError` on the first expression that fails, naming the sheet and the addresses.
   *   Choose it when a dropped formula makes the result wrong rather than merely poorer — a pipeline that recalculates
   *   the workbook would otherwise silently inherit a constant where a computation belonged.
   *
   * Ignored when reading XLSX, where a formula is stored as the text a caller would get back anyway.
   */
  readonly formulas?: "preserve" | "cached" | "error";
}

/**
 * Options for `readWithDiagnostics`.
 *
 * `unsupported` is deliberately **not** among them. That function's whole purpose is to hand back the
 * report instead of turning it into a rejection, so accepting a policy it then ignores would be a
 * silently inert option — and one whose obvious reading (`"error"` should throw) is the opposite of what
 * would happen. Omitting it from the type says so at the call site rather than in a paragraph.
 */
export type WorkbookDiagnosticReadOptions = Omit<WorkbookReadOptions, "unsupported">;

/** Options common to writing either format. */
export interface WorkbookWriteOptions extends XlsxWriteOptions {
  /** Format to write. Defaults to `"xlsx"`, or to the extension for `writeFile`. */
  readonly format?: WorkbookFormat;
  /**
   * What to do about content the XLSB writer cannot express.
   *
   * `"error"` (the default) refuses the write and names what would be lost, on the thrown error's
   * `items`. `"ignore"` writes the workbook without it. Ignored when writing XLSX, which expresses
   * everything this library models.
   */
  readonly unsupported?: "error" | "ignore";
  /**
   * Encrypt the workbook with this password (1 to 255 characters) — Agile Encryption in an OLE compound file, what
   * Excel writes for "Encrypt with Password", for XLSX and XLSB alike; `Workbook.read(bytes, { password })` reads it
   * back. The same option Word's `Document.toBuffer` takes. The encryption code is loaded only for a write given one, and
   * the package is assembled in memory before it is encrypted — `writeStream` and `toStream` included. Not
   * `Workbook.protect`'s password, which only locks the workbook's structure.
   */
  readonly password?: string;
  /**
   * How {@link password} encrypts: key size, hash algorithm, spin count. Default: AES-256, SHA-512 and 100,000
   * iterations, as Office writes. Only with `password`.
   */
  readonly encryption?: EncryptOptions;
}

/**
 * What a read recovered, and what it could not.
 *
 * Returned by `readWithDiagnostics` for the case `unsupported` cannot serve: a caller who wants the
 * workbook *and* the report. `"error"` gives the report only by refusing, and the default gives the
 * workbook only by staying silent — so a tool that reads a file and then tells its user "these cells
 * could not be recovered" had no way to be written.
 */
export interface WorkbookReadReport {
  /** The workbook, already populated. */
  readonly workbook: WorkbookData;
  /**
   * Content that did not reach the workbook, as `Sheet!A1: reason` or `Sheet: reason`.
   *
   * Empty for an XLSX read, which has no equivalent channel — not because nothing is ever lost there, but
   * because that reader does not collect one. Saying so is better than implying a guarantee. A setting whose value is
   * outside its vocabulary is not lost by either reader; see `unrecognizedValues`.
   */
  readonly lost: readonly string[];
  /**
   * Cell records recognised but not decodable, counted by record name.
   *
   * The same losses `lost` describes in prose, in the shape a program can act on. A converter reporting
   * "1,400 cells use `BrtShortReal`" should not have to parse a sentence to find that out, and the
   * sentence deliberately *aggregates* — one line per sheet — so the numbers are not all in it.
   */
  readonly unreadRecords: ReadonlyMap<string, number>;
  /** Formula expressions that could not be decoded, as `Sheet!A1`. The cached value survived. */
  readonly undecodedFormulas: readonly string[];
  /** Cells deferring to a shared formula, which this reader does not resolve, as `Sheet!A1`. */
  readonly sharedFormulaCells: readonly string[];
  /**
   * Record ids this library has no name for, counted by id.
   *
   * Deliberately *not* part of `lost`, and this is the field that makes the distinction usable. A record
   * from a newer schema is not missing content — every workbook in the reference corpus has some — so
   * counting them as losses would make `unsupported: "error"` reject ordinary files. A caller who wants
   * to know anyway can look here.
   */
  readonly unknownRecords: ReadonlyMap<number, number>;
  /**
   * Settings the file states with a value outside what the setting takes, one entry each:
   * `{ sheet: "S", path: "pageSetup.orientation", value: "sideways" }` — the shape the
   * `Workbook.discardUnrecognizedValues` filter receives, and the `{ path, value }` Word reports.
   *
   * Not losses, and that is why they are not in `lost`: each is kept as read — its field reads as absent — and written
   * back unchanged by an XLSX write while the field is absent; `Workbook.discardUnrecognizedValues` drops them. Both
   * readers fill it. An entry for a value the
   * format does allow means this library's vocabulary for it is too narrow, which is worth reporting.
   */
  readonly unrecognizedValues: readonly UnrecognizedValue[];
}

export type { UnrecognizedValue } from "@excel/core/unrecognized-values";
export type { EncryptOptions as WorkbookEncryptOptions } from "@archive/office-crypto/agile";

/**
 * Whether the first argument of a read is the workbook to read *into* — the deprecated form — rather than the input.
 *
 * Decided structurally, because neither side has a brand to test: a workbook handle is a plain object that carries
 * its sheet list, and no input a read accepts does — bytes are typed arrays or `ArrayBuffer`s, a path or base64 text
 * is a string, and a stream is an object with no `_worksheets`. Typed arrays are excluded explicitly so a caller's
 * `Uint8Array` subclass with arbitrary own properties can never be taken for a workbook.
 */
export function isReadTarget(value: unknown): value is WorkbookData {
  return (
    typeof value === "object" &&
    value !== null &&
    !ArrayBuffer.isView(value) &&
    !(value instanceof ArrayBuffer) &&
    Array.isArray((value as { _worksheets?: unknown })._worksheets)
  );
}

/** Split the create options a new-workbook read accepts from the read's own. */
export function splitCreateOptions<T extends object>(
  options: (T & WorkbookCreateOptions) | undefined
): { create: WorkbookCreateOptions | undefined; read: T | undefined } {
  // Anything but an object is passed through for the read's own option check to refuse.
  if (typeof options !== "object" || options === null || !("formulaSyntaxProbe" in options)) {
    return { create: undefined, read: options };
  }
  const { formulaSyntaxProbe, ...read } = options;
  return { create: { formulaSyntaxProbe }, read: read as T };
}

/** A limit: a whole number of at least `min`, or `Infinity` to lift it. `NaN` and `0` used to disable a guard. */
function limit(min: number) {
  const expected = `an integer >= ${min}, or Infinity`;
  return s
    .union([s.literal(Infinity), s.number({ integer: true, min, expected })], { expected })
    .optional();
}

const format = () => s.literal("xlsx", "xlsb").optional();
const flag = () => s.boolean().optional();

const zipReadLimitsSchema = s.object({
  maxEntrySize: limit(0),
  maxTotalUncompressedSize: limit(0),
  maxEntries: limit(0),
  checkCrc32: flag()
});
const _zipReadLimitsMatch: SchemaMatches<typeof zipReadLimitsSchema, XlsxZipReadLimits> = true;
void _zipReadLimitsMatch;

const xlsxReadShape = {
  base64: flag(),
  maxRows: limit(1),
  maxCols: limit(1),
  // A string matched every node name containing it: `"sheetData"` dropped every cell of every sheet.
  ignoreNodes: s.array(s.string({ nonEmpty: true })).optional(),
  zip: zipReadLimitsSchema.optional(),
  password: s.string().optional()
};
const xlsxReadOptionsSchema = s.object(xlsxReadShape);
const _xlsxReadOptionsMatch: SchemaMatches<typeof xlsxReadOptionsSchema, XlsxReadOptions> = true;
void _xlsxReadOptionsMatch;

const diagnosticReadOptionsSchema = s.extend(
  xlsxReadOptionsSchema,
  {
    format: format(),
    blankCells: s.literal("keep", "collapse").optional(),
    formulas: s.literal("preserve", "cached", "error").optional()
  },
  { readonly: true }
);
const _diagnosticReadOptionsMatch: SchemaMatches<
  typeof diagnosticReadOptionsSchema,
  WorkbookDiagnosticReadOptions
> = true;
void _diagnosticReadOptionsMatch;

const readOptionsSchema = s.extend(
  diagnosticReadOptionsSchema,
  { unsupported: s.literal("error", "ignore").optional() },
  { readonly: true }
);
const _readOptionsMatch: SchemaMatches<typeof readOptionsSchema, WorkbookReadOptions> = true;
void _readOptionsMatch;

/** `Workbook.create`'s options, which a read that returns a new workbook takes too. `null` has always meant none. */

const zipWriterOptionsSchema = s.object({
  // Past 0–9 the deflater refuses the level, and the refusal used to leave the write pending forever.
  level: s.number({ integer: true, min: 0, max: 9 }).optional(),
  modTime: s
    .custom<Date>(v => v instanceof Date && !Number.isNaN(v.getTime()), "a valid Date")
    .optional(),
  timestamps: s.literal("dos", "dos+utc", "dos+utc+ntfs").optional()
});
const _zipWriterOptionsMatch: SchemaMatches<typeof zipWriterOptionsSchema, ZipWriterOptions> = true;
void _zipWriterOptionsMatch;

const xlsxWriteShape = {
  zip: zipWriterOptionsSchema.optional(),
  useSharedStrings: flag(),
  useStyles: flag(),
  templateMode: s.literal("preserve", "strict").optional(),
  strictTemplateMode: flag(),
  validate: flag()
};
const xlsxWriteOptionsSchema = s.object(xlsxWriteShape);
const _xlsxWriteOptionsMatch: SchemaMatches<typeof xlsxWriteOptionsSchema, XlsxWriteOptions> = true;
void _xlsxWriteOptionsMatch;

const encryptOptionsSchema = s.object(
  {
    keyBits: s.literal(128, 192, 256).optional(),
    hashAlgorithm: s
      .literal("SHA-1", "SHA-256", "SHA-384", "SHA-512", "SHA1", "SHA256", "SHA384", "SHA512")
      .optional(),
    spinCount: s.number({ integer: true, min: 0, max: 10_000_000 }).optional()
  },
  { readonly: true }
);
const _encryptOptionsMatch: SchemaMatches<typeof encryptOptionsSchema, EncryptOptions> = true;
void _encryptOptionsMatch;

const xlsxStreamOptionsSchema = s.object({
  ...xlsxWriteShape,
  highWaterMark: s.number({ integer: true, min: 1 }).optional()
});
const _xlsxStreamOptionsMatch: SchemaMatches<typeof xlsxStreamOptionsSchema, XlsxStreamOptions> =
  true;
void _xlsxStreamOptionsMatch;

const writeOptionsSchema = s.extend(
  xlsxWriteOptionsSchema,
  {
    format: format(),
    unsupported: s.literal("error", "ignore").optional(),
    // MS-OFFCRYPTO caps a password at 255 characters, and an empty one would be the absence of one.
    password: s
      .string({ expected: "a string of 1 to 255 characters" })
      .refine(v => v.length > 0 && v.length <= 255, "a string of 1 to 255 characters")
      .optional(),
    encryption: encryptOptionsSchema.optional()
  },
  { readonly: true }
);
const streamOptionsSchema = s.extend(writeOptionsSchema, {
  highWaterMark: s.number({ integer: true, min: 1 }).optional()
});
const _writeOptionsMatch: SchemaMatches<typeof writeOptionsSchema, WorkbookWriteOptions> = true;
void _writeOptionsMatch;

/**
 * A read's or a write's options: an object, if given, whose `format` is one this library reads and writes. A format it
 * did not know was ignored — `toBuffer(wb, { format: "csv" })` wrote XLSX — and options that were not an object reached
 * an `in` test as a raw `TypeError`. `null` is taken as no options, as it always was.
 */
const createAndReadOptionsSchema = s.all(createOptionsSchema, readOptionsSchema);

export function checkIoOptions(
  where: string,
  options: unknown,
  direction: "read" | "write" | "stream"
): void {
  if (options === null || options === undefined) {
    return;
  }
  if (direction === "read") {
    checkSchema(where, "options", createAndReadOptionsSchema, options);
    return;
  }
  checkSchema(
    where,
    "options",
    direction === "stream" ? streamOptionsSchema : writeOptionsSchema,
    options
  );
  const { password, encryption } = options as WorkbookWriteOptions;
  if (encryption !== undefined && password === undefined) {
    throw argumentFailure(
      where,
      "options.encryption",
      "set only with options.password",
      encryption
    );
  }
}

/** Whether `value` is a `Blob` (a `File` included) — recognised structurally, as a Blob from another realm is one. */
export function isBlob(value: unknown): value is Blob {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as Blob).arrayBuffer === "function" &&
    typeof (value as Blob).size === "number" &&
    typeof (value as Blob).type === "string"
  );
}

/** `data` with a `Blob` read to its bytes — what the async whole-workbook readers take besides bytes and base64 text. */
export async function blobBytes<T>(data: T | Blob): Promise<T | Uint8Array> {
  return isBlob(data) ? new Uint8Array(await data.arrayBuffer()) : data;
}

/**
 * The input of `read` / `readWithDiagnostics`: bytes — a `Uint8Array` or any other `ArrayBufferView`, an
 * `ArrayBuffer` from any realm, a `Blob` (read first, by {@link blobBytes}) — or base64 text with `{ base64: true }`. Anything else (`42`, `null`, `{}`, text without
 * the flag) is an argument error, as it is for `readStream`; it used to reach the ZIP loader and come back as an
 * `ExcelFileError` about "the loaded zip file". Real bytes that are not a workbook stay an `ExcelFileError`.
 */
export function checkReadData(where: string, data: unknown, base64: unknown): void {
  if (ArrayBuffer.isView(data) || isArrayBufferLike(data)) {
    return;
  }
  if (typeof data === "string" && base64 === true) {
    return;
  }
  const got =
    typeof data === "string"
      ? "a string without { base64: true }"
      : data === null
        ? "null"
        : Array.isArray(data)
          ? "an array"
          : typeof data;
  throw new ExcelArgumentError(
    "data",
    `${where}: data must be bytes (Uint8Array, ArrayBuffer, a typed-array view, a Blob) or a base64 string with { base64: true }, got ${got}`
  );
}

/**
 * The input of `readStream`: something the read can pump — a Node stream, a web `ReadableStream`, or an (async)
 * iterable of chunks. Anything else used to reach the pipeline, whose refusal was never observed: the read waited for
 * entries that could not come, and the error surfaced as an unhandled rejection.
 */
export function checkReadStreamInput(where: string, value: unknown): void {
  const candidate = value as {
    pipe?: unknown;
    getReader?: unknown;
    [Symbol.asyncIterator]?: unknown;
    [Symbol.iterator]?: unknown;
  } | null;
  if (
    typeof value === "object" &&
    candidate !== null &&
    // Bytes are iterable — of numbers, not chunks — so they would pass for a stream and fail inside it.
    !ArrayBuffer.isView(value) &&
    (typeof candidate.pipe === "function" ||
      typeof candidate.getReader === "function" ||
      typeof candidate[Symbol.asyncIterator] === "function" ||
      typeof candidate[Symbol.iterator] === "function")
  ) {
    return;
  }
  argumentChecks(where).fail(
    "stream",
    ArrayBuffer.isView(value)
      ? "a readable stream, a ReadableStream or an iterable of byte chunks (pass bytes to Workbook.read)"
      : "a readable stream, a ReadableStream or an iterable of byte chunks",
    value
  );
}

/**
 * `readWithDiagnostics`, once, for both platforms.
 *
 * **The body was 26 identical lines in `xlsx-io.ts` and `xlsx-io.browser.ts`.** This file already exists because the
 * *types* were declared twice and drifted — its own header says so — and the function that produces them was duplicated
 * beside them. Adding a field to `WorkbookReadReport` meant editing two bodies, and missing one would leave a platform
 * returning a report short a field without any type error, because the type came from here.
 *
 * `read` is injected rather than imported: it is the one genuinely platform-specific piece, since Node accepts a file
 * path and a `Buffer` where a browser does not. Everything else — the format resolution, the XLSB branch, the empty
 * report an XLSX read yields — is identical and now exists once.
 *
 * `parseXlsbPackage` and `commitXlsbRead` used to be injected too, and should not have been: both variants passed the
 * same pair, so the injection only spread a static import of the binary reader across both of them. They are loaded
 * where they are used now.
 */
export async function readWorkbookWithDiagnostics(
  wb: WorkbookData,
  input: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  options: WorkbookDiagnosticReadOptions | undefined,
  platform: {
    readonly read: (
      wb: WorkbookData,
      data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
      options?: WorkbookDiagnosticReadOptions
    ) => Promise<WorkbookData>;
    readonly normalizeBytes: (
      data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
      base64?: boolean
    ) => Uint8Array | undefined;
    readonly resolveReadFormat: (
      bytes: Uint8Array,
      format?: WorkbookFormat,
      zip?: XlsxZipReadLimits
    ) => WorkbookFormat;
  }
): Promise<WorkbookReadReport> {
  checkIoOptions("Workbook.readWithDiagnostics", options, "read");
  const data = await blobBytes(input);
  checkReadData("Workbook.readWithDiagnostics", data, options?.base64);
  const raw = platform.normalizeBytes(data, options?.base64);
  // An encrypted workbook is decrypted first, so its package's format is the one detected.
  const bytes = raw && (await openWorkbookPackage(raw, workbookDecryption(options)));
  if (
    bytes !== undefined &&
    platform.resolveReadFormat(bytes, options?.format, options?.zip) === "xlsb"
  ) {
    // Loaded here rather than injected by the caller, and on demand rather than statically. The two
    // platform variants passed in the *same* two functions — they are not platform-specific — and a
    // static import of them in either made the whole binary reader eager for every `Workbook.read`
    // consumer, XLSX-only ones included, and defeated the streaming reader's own `await import()` of
    // the same modules. This branch is the only place they are used, and it is already async.
    const { commitXlsbRead, parseXlsbPackage } = await import("@excel/xlsb/read/package");
    // `blankCells` reaches here too: `readWithDiagnostics` is the same read with the report handed back instead of
    // thrown, so a caller inspecting a large formatted sheet needs the same policy the plain read offers.
    const parsed = await parseXlsbPackage(bytes, "<buffer>", xlsbReadOptions(options));
    commitXlsbRead(wb, parsed);
    wb.sourceFilePath = undefined;
    return { workbook: wb, ...parsed.diagnostics, unrecognizedValues: keptSettingsReport(wb) };
  }
  const workbook = await platform.read(wb, bytes ?? data, options);
  return {
    // **An XLSX read reports no losses because that reader collects nothing, not because nothing can be lost.** So an empty
    // report here means "not measured", which `WorkbookReadReport.lost` says in prose.
    //
    // It does *not* mean the reader truncates behind the caller's back: `maxRows` and `maxCols` throw
    // `ExcelLimitExceededError` rather than skipping, which their own documentation used to deny and their tests have
    // always required.
    workbook,
    lost: [],
    unreadRecords: new Map(),
    undecodedFormulas: [],
    sharedFormulaCells: [],
    unknownRecords: new Map(),
    unrecognizedValues: keptSettingsReport(workbook)
  };
}

/**
 * What `Workbook.readStream` reads: a Node `Readable`, a web `ReadableStream`, or any (async) iterable of the
 * package's byte chunks.
 */
export type WorkbookStreamInput =
  | AsyncIterable<Uint8Array | string>
  | ReadableStream<Uint8Array>
  | Iterable<Uint8Array>
  | XlsxParseStream;

/**
 * `readStream`, once, for both platforms: an XLSX package is read as it streams; an XLSB one — named by
 * `options.format`, since a stream cannot be looked at before it is read — is collected and read as `Workbook.read`
 * reads its bytes. `readStream` used to take XLSX only, and on an XLSB stream its error told the caller to pass a
 * `format: "xlsb"` it then ignored.
 */
export async function readWorkbookStream(
  wb: WorkbookData,
  stream: unknown,
  options: WorkbookReadOptions | undefined,
  platform: {
    readonly read: (
      wb: WorkbookData,
      data: Uint8Array,
      options?: WorkbookReadOptions
    ) => Promise<WorkbookData>;
  }
): Promise<WorkbookData> {
  checkReadStreamInput("Workbook.readStream", stream);
  checkIoOptions("Workbook.readStream", options, "read");
  if (options?.format === "xlsb") {
    const { streamToUint8Array } = await import("@stream/core/stream-helpers");
    const bytes = await streamToUint8Array(
      stream as AsyncIterable<Uint8Array> | ReadableStream<Uint8Array> | Iterable<Uint8Array>
    );
    return platform.read(wb, bytes, options);
  }
  const { readXlsxStreamInto } = await import("@excel/xlsx/read/stream");
  return readXlsxStreamInto(wb, stream as never, options);
}

/** The decryption a buffered read's options ask for, its errors naming the format read. */
export function workbookDecryption(
  options: WorkbookReadOptions | undefined,
  source?: string
): WorkbookDecryption {
  return {
    ...decryptionOf(options, source),
    format: options?.format === "xlsb" ? "XLSB" : "XLSX"
  };
}
