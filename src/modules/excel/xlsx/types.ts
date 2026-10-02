/**
 * Option types and stream contracts shared by the XLSX reader (`read/`) and writer (`write/`).
 *
 * Types only, so either side can name them without reaching the other.
 */
import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
import type { XlsxEmitterLike, XlsxWritable } from "@excel/core/xlsx-io-types";
import type { IEventEmitter } from "@stream";

export type StreamListener = Parameters<IEventEmitter["on"]>[1];

/**
 * Event-emitter basics shared by every stream shape here. Aliased to the public
 * IO contract so the serializer and `Workbook.writeStream` cannot drift apart.
 */
type EmitterLike = XlsxEmitterLike;

export interface IParseStream extends EmitterLike {
  pipe(dest: any): any;
  [Symbol.asyncIterator]?: () => AsyncIterator<Uint8Array | string>;
}

export interface IZipWriter extends EmitterLike {
  append(data: string | Uint8Array, options: { name: string; base64?: boolean }): void;
  /**
   * Create a streaming entry: write chunks incrementally, then call end().
   * The entry's compressed output is tracked internally, so a later
   * `waitForDrain()` observes it even when compression is asynchronous.
   */
  /**
   * `Uint8Array` as well as `string`, because a binary part cannot go through the string form: the entry
   * encodes text as UTF-8, so every byte above 0x7F would become two. XML parts still pass strings.
   */
  createEntry(name: string): { write(chunk: Uint8Array | string): void; end(): void };
  pipe(stream: XlsxWritable): void;
  finalize(): void;
  /** Wait for downstream backpressure to clear. Resolves immediately if no backpressure. */
  waitForDrain(): Promise<void>;
}

/**
 * Options for reading (loading) an XLSX workbook.
 *
 * The option set is closed. It used to carry an `[key: string]: unknown` index
 * signature "for forward compatibility", which in practice meant a misspelled or
 * unsupported option type-checked and was then silently ignored — the failure
 * mode a typed options bag exists to prevent. A new option is a new declared
 * field here.
 */
export interface XlsxReadOptions {
  /**
   * When the input to `load()` is a string, interpret it as a base64-encoded
   * zip archive instead of a binary buffer. Defaults to `false`.
   */
  base64?: boolean;
  /**
   * Refuse a worksheet with more rows than this, by throwing `MaxItemsExceededError`.
   *
   * **A guard, not a preview.** This used to say "rows beyond this limit are silently skipped. Useful for previewing
   * very large sheets without loading everything into memory" — which is the opposite of what happens and of what the
   * tests require: `ListXform.parseClose` throws, and the cases covering it are named "should bail out" and "should fail
   * fast on a huge file". A caller who followed the old description got an exception instead of a truncated read.
   *
   * The distinction matters beyond wording: a limit that *fails* is a way to decline hostile input, while one that
   * truncates is a way to lose data quietly. This is the first, and there is no option for the second.
   */
  maxRows?: number;
  /**
   * Refuse a worksheet with more columns than this, by throwing `MaxItemsExceededError`.
   *
   * A guard, exactly as {@link maxRows} is — and its description was wrong in the same way.
   */
  maxCols?: number;
  /**
   * List of worksheet XML node names to skip while parsing (e.g.
   * `"dataValidations"`, `"conditionalFormatting"`). Use for workbooks that
   * contain corrupted or unsupported elements you want to ignore.
   */
  ignoreNodes?: string[];
}

export interface ZipWriterOptions {
  level?: number;
  /** ZIP entry modification time (optional). If omitted, defaults to current time. */
  modTime?: Date;
  /** Timestamp writing strategy for ZIP entry metadata (optional). */
  timestamps?: ZipTimestampMode;
}

export type XlsxTemplateMode = "preserve" | "strict";

/**
 * Options for writing an XLSX workbook.
 *
 * The option set is closed — see {@link XlsxReadOptions} for why. This matters
 * more here than on the read side: a write option that is accepted and dropped
 * produces a file that differs from what the caller asked for, with nothing
 * anywhere reporting it.
 */
export interface XlsxWriteOptions {
  /** ZIP archive options (compression level, timestamps, ...). */
  zip?: ZipWriterOptions;
  /**
   * Use a shared-string table for cell text values. Defaults to `true`.
   * Set to `false` to write string values inline (larger file, but streams
   * better for very large sheets).
   */
  useSharedStrings?: boolean;
  /**
   * Emit style definitions (fonts, fills, borders, number formats, …).
   * Defaults to `true`. Set to `false` to skip style blocks for maximum
   * compatibility with minimal readers.
   */
  useStyles?: boolean;
  /**
   * Template fidelity strategy for loaded workbooks. The default `"preserve"`
   * byte-preserves clean chart parts and may structurally re-render edited parts
   * when no safe raw XML patch is available. `"strict"` fails the write instead
   * of re-rendering any edited loaded chart/chartEx part that cannot be patched
   * in-place, preventing silent loss of unknown template XML.
   *
   * When a strict write fails, the thrown error enumerates the unrecognised
   * `c15:`/`cx14:` extension paths the parser observed, so authors can decide
   * between relaxing the mode or reshaping the mutation into a patch-friendly
   * path. To inspect those paths *before* writing (for example to decide
   * whether to opt into strict mode at all), use {@link Chart.unknownElements}
   * on each chart of interest.
   */
  templateMode?: XlsxTemplateMode;
  /**
   * Convenience alias for `templateMode: "strict"`. Strict mode refuses to
   * silently drop vendor-extension XML (`c15:…`, `cx14:…`) that the
   * structured parser does not understand. Use this when round-tripping
   * Excel-authored template files where preserving exotic extension
   * elements matters more than the ability to re-render modified charts.
   */
  strictTemplateMode?: boolean;
  /**
   * Run the OOXML self-check on the produced bytes and `console.warn`
   * every detected problem. Only `Workbook.toBuffer` honours this
   * flag — `write(stream)` / `writeFile` cannot post-validate because
   * their output is streamed to the caller.
   *
   * Default resolution:
   *   - In Node.js with `NODE_ENV !== "production"` and NOT running
   *     under vitest (`process.env.VITEST !== "true"`): `true`.
   *   - In production, in the browser, or under vitest: `false`.
   *
   * The vitest carve-out exists because running validation on every
   * `Workbook.toBuffer` call inflates fixture `beforeAll` hooks that build
   * hundreds of workbooks by ~50 seconds on typical hardware. Tests
   * that want validation use {@link expectValidXlsx} directly.
   *
   * Pass `true` to force validation (even in production or under
   * vitest), or `false` to suppress it (even in development). The
   * self-check never throws: a failed validation becomes a warning so
   * writers that intentionally produce non-conformant xlsx for
   * testing keep working.
   *
   * **It dominates a development-mode write, so benchmark with
   * `NODE_ENV=production`.** Measured on a 100-column × 200-row table:
   * 38 ms with the check off against 124 ms with it on — 3.3×. The cost
   * is not waste that could be optimised away; the check parses each
   * part once into a DOM that the structural checks then share. But it
   * does mean a profile taken in development attributes time to the
   * writer that production never spends, which is how a reported
   * "13% in style registration" turned out to be measured against a
   * denominator inflated by validation.
   *
   * @see OoxmlValidationReport
   */
  validate?: boolean;
}

export type XlsxOptions = XlsxReadOptions & XlsxWriteOptions;

export interface WorkbookMediaLike {
  type: string;
  extension: string;
  name?: string;
  filename?: string;
  buffer?: Uint8Array;
  base64?: string;
  /** External link target — when set, the image is referenced, not embedded. */
  link?: string;
}

export interface MediaModel {
  media: WorkbookMediaLike[];
}

export interface ZipEntryLike {
  name: string;
  // Mirrors the `type` field on streaming ZipEntry. Symlink detection in
  // streaming parsers is best-effort; for XLSX (which contains no symlinks)
  // this is effectively "Directory" | "File" at runtime.
  type: "Directory" | "File" | "Symlink";
  stream: AsyncIterable<Uint8Array | string>;
  drain: () => Promise<void>;
}
