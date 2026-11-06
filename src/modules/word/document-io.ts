/**
 * DOCX Module - Document IO
 *
 * IO operations: package/unpackage to bytes, base64, Flat OPC,
 * plus the public Patch and Template APIs that combine read/patch/write.
 */

import type { EncryptOptions } from "@archive/office-crypto/agile";
import { unzipZip } from "@archive/read-archive";
import { asPlainBytes } from "@utils/binary";
import { readFileBytes, writeFileBytes } from "@utils/fs";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
import {
  checkBytes,
  checkHandle,
  checkPackageOptions,
  invalidArgument,
  packageOptionsShape
} from "@word/core/argument-checks";
import { bytesToBase64 } from "@word/core/internal-utils";
import { checkSchema, model, w } from "@word/core/option-schema";
import { DocxFileError } from "@word/errors";
import type { PatchOperation } from "@word/patcher";
import { applyPatchesToDocument, checkPatches } from "@word/patcher";
import { checkReadOptions, readDocx } from "@word/reader/docx-reader";
import type { ReadDocxOptions } from "@word/reader/docx-reader";
import { checkWriteEncryption, encryptIfRequested } from "@word/security/package-encryption";
import { checkSecurityPolicy } from "@word/security/policy";
import type { DocxSecurityPolicy } from "@word/security/policy";
import { fillData, fillDocument, fillTemplate } from "@word/template/template-engine";
import type { LegacyTemplateOptions, TemplateOptions } from "@word/template/template-engine";
import type { DocxDocument } from "@word/types";
import type { DocxWriteOptions } from "@word/writer/docx-packager";
import { packageDocx } from "@word/writer/docx-packager";

// Re-export patch types for backward compatibility
export type { PatchContent, PatchOperation } from "@word/patcher";

// =============================================================================
// Document IO (toBuffer / toBase64)
// =============================================================================

/**
 * Package a DocxDocument model to DOCX bytes.
 *
 * With `{ password }` the package is encrypted — Agile Encryption in a compound file, what Word writes for
 * "Encrypt with Password" — and `Document.read(bytes, { password })` reads it back; `encryption` tunes it (key size,
 * hash, spin count). Without one it is written unencrypted, whatever the document was read from. The encryption
 * code is loaded only for a write that asks for it.
 *
 * ```ts
 * const encrypted = await Document.toBuffer(doc, { password: "secret" });
 * const again = await Document.read(encrypted, { password: "secret" });
 * ```
 *
 * @throws {DocxDuplicateIdError} (as a rejection) If two footnotes, endnotes,
 * comments, numbering definitions or builder images share an id — see
 * `DocxDocument`.
 */
export async function toBuffer(
  doc: DocxDocument,
  options?: DocxWriteOptions
): Promise<Uint8Array<ArrayBuffer>> {
  checkHandle(doc, "document", "doc", "Document.toBuffer");
  checkWriteOptions(options, "Document.toBuffer");
  return asPlainBytes(await encryptIfRequested(await packageDocx(doc, options), options));
}

/** Package a DocxDocument model to base64 string — encrypted with `{ password }`, as {@link toBuffer}. */
export async function toBase64(doc: DocxDocument, options?: DocxWriteOptions): Promise<string> {
  checkHandle(doc, "document", "doc", "Document.toBase64");
  checkWriteOptions(options, "Document.toBase64");
  const bytes = await encryptIfRequested(await packageDocx(doc, options), options);
  return bytesToBase64(bytes);
}

// =============================================================================
// File-path IO (Node only)
// =============================================================================
//
// Convenience wrappers that read/write a DOCX directly from/to the filesystem,
// mirroring the Excel module's `Workbook.readFile` / `Workbook.writeFile`. The
// browser build of the `Document` namespace (`surface/document.browser.ts`) does not publish
// them, as Excel's and Archive's browser entries publish no file-path API.

/**
 * Read a DOCX file from a filesystem path and parse it into a `DocxDocument`.
 *
 * Node-only. In the browser, read the bytes yourself and use `Document.read`.
 *
 * @param path - Filesystem path to the `.docx` file.
 * @param options - Read options (password, security policy).
 * @throws {DocxArgumentError} If `path` is not a non-empty string.
 * @throws {DocxFileError} If the file system will not let the file be read
 *   (`ENOENT`, `EACCES`, …); the system's error is its `cause`.
 */
export async function readFile(path: string, options?: ReadDocxOptions): Promise<DocxDocument> {
  checkPath(path, "Document.readFile");
  checkReadOptions(options, "Document.readFile");
  let bytes: Uint8Array;
  try {
    bytes = await readFileBytes(path);
  } catch (error) {
    throw fileSystemError(path, "read", error);
  }
  return readDocx(bytes, options);
}

/**
 * Package a `DocxDocument` and write it to a filesystem path.
 *
 * Node-only. In the browser, use {@link toBuffer} and persist the bytes yourself.
 *
 * @param doc - The document model to serialize.
 * @param path - Destination filesystem path.
 * @param options - Packaging options, as {@link toBuffer} takes them (`password` encrypts the file).
 * @throws {DocxArgumentError} If `path` is not a non-empty string.
 * @throws {DocxFileError} If the file system will not let the file be written;
 *   the system's error is its `cause`.
 */
export async function writeFile(
  doc: DocxDocument,
  path: string,
  options?: DocxWriteOptions
): Promise<void> {
  checkHandle(doc, "document", "doc", "Document.writeFile");
  checkPath(path, "Document.writeFile");
  checkWriteOptions(options, "Document.writeFile");
  const bytes = await encryptIfRequested(await packageDocx(doc, options), options);
  try {
    await writeFileBytes(path, bytes);
  } catch (error) {
    throw fileSystemError(path, "write", error);
  }
}

/** The `DocxFileError` for a file the system would not let us read or write, with its error as the cause. */
function fileSystemError(path: string, operation: "read" | "write", error: unknown): DocxFileError {
  const code = (error as { code?: unknown } | null)?.code;
  const detail =
    code === "ENOENT" ? "File not found" : error instanceof Error ? error.message : String(error);
  return new DocxFileError(path, operation, detail, { cause: error });
}

const patchOptionsShape = {
  ...packageOptionsShape,
  securityPolicy: model<DocxSecurityPolicy>()
};

const writeOptionsSchema = w.record({
  ...patchOptionsShape,
  // MS-OFFCRYPTO §2.3.4.10 caps a password at 255 characters, and an empty one would be the
  // absence of one.
  password: s
    .custom<string>(
      v => typeof v === "string" && v.length > 0 && v.length <= 255,
      "a string of 1 to 255 characters"
    )
    .optional(),
  encryption: model<EncryptOptions>()
});
const _writeOptionsMatch: SchemaMatches<typeof writeOptionsSchema, DocxWriteOptions> = true;
void _writeOptionsMatch;

const patchOptionsSchema = w.record(patchOptionsShape);
const _patchOptionsMatch: SchemaMatches<typeof patchOptionsSchema, PatchOptions> = true;
void _patchOptionsMatch;

const toFlatOpcOptionsSchema = w.record({
  compressionLevel: packageOptionsShape.compressionLevel
});
const _toFlatOpcOptionsMatch: SchemaMatches<typeof toFlatOpcOptionsSchema, ToFlatOpcOptions> = true;
void _toFlatOpcOptionsMatch;

/** The security policy of a write's options; `null` has always read as the default policy. */
function checkWritePolicy(options: { securityPolicy?: unknown } | undefined, where: string): void {
  if (options?.securityPolicy != null) {
    checkSecurityPolicy(options.securityPolicy, "options.securityPolicy", where);
  }
}

/**
 * The write options of `Document.toBuffer` and its siblings, checked in the name of
 * `where` — the security policy included, so a bad one is reported as
 * `options.securityPolicy.…` of the call, not by the packager.
 */
function checkWriteOptions(options: unknown, where: string): void {
  checkSchema(writeOptionsSchema.optional(), options, where);
  checkWritePolicy(options, where);
  checkWriteEncryption(options, where);
}

/** The write side of a `Package` function's options (`PatchOptions`), checked in the name of `where`. */
function checkPatchOptions(options: unknown, where: string): void {
  checkSchema(patchOptionsSchema.optional(), options, where);
  checkWritePolicy(options, where);
}

/** A file path as `Workbook.readFile` takes one: a non-empty string. */
function checkPath(path: unknown, where: string): void {
  if (typeof path !== "string" || path === "") {
    invalidArgument(where, "path", "a non-empty string", path);
  }
}

// =============================================================================
// Patcher / Template Fill API
// =============================================================================

/**
 * Options for `Package.patch` / `Package.patchTemplate`: how the result is
 * written, as `Document.toBuffer` takes them.
 */
export interface PatchOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
  /** Timestamp for every entry of the written package, as in `Document.toBuffer`. */
  readonly modTime?: Date;
  /** Security policy for writing (raw-XML policy, preserved parts), as in `Document.toBuffer`. */
  readonly securityPolicy?: DocxSecurityPolicy;
}

/**
 * A compiled template that can be reused for multiple patch operations. Treat it as opaque: pass
 * it to `Package.patchTemplate`.
 */
export interface CompiledTemplate {
  /**
   * @deprecated The parsed document, held for `Package.patchTemplate`; not part of the contract. Read
   * the template with `Document.read` instead. Leaves the public type in 1.0.
   */
  readonly _doc: DocxDocument;
}

/**
 * Compile a DOCX template for reuse with multiple data sets.
 *
 * Parsing the ZIP and XML is expensive. If you need to patch the same template
 * multiple times with different data, compile it once and reuse.
 *
 * @param buffer - The source DOCX template file.
 * @param options - How to read it, as `Document.read` takes them (`password`, `securityPolicy`).
 * @returns A compiled template handle.
 */
export async function compileTemplate(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  options?: ReadDocxOptions
): Promise<CompiledTemplate> {
  const where = "Package.compileTemplate";
  const doc = await readPackage(buffer, options, where);
  return { _doc: doc };
}

/**
 * Apply patches to a pre-compiled template.
 *
 * Much faster than `patchDocument` when applying to the same template repeatedly,
 * since ZIP/XML parsing is skipped.
 *
 * @param template - A compiled template from `compileTemplate`.
 * @param patches - Array of patch operations to apply.
 * @param options - How the result is written, as `Document.toBuffer` takes them.
 * @returns New DOCX file as a Uint8Array.
 */
export async function patchTemplate(
  template: CompiledTemplate,
  patches: readonly PatchOperation[],
  options?: PatchOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const doc = (template as { _doc?: unknown } | null | undefined)?._doc;
  if (typeof doc !== "object" || doc === null || !Array.isArray((doc as { body?: unknown }).body)) {
    invalidArgument(
      "Package.patchTemplate",
      "template",
      "a CompiledTemplate (from Package.compileTemplate)",
      template
    );
  }
  checkPatches(patches, "Package.patchTemplate");
  checkPatchOptions(options, "Package.patchTemplate");
  // Deep clone the document to avoid mutating the cached template
  const patched = applyPatchesToDocument(structuredClone(template._doc), patches);
  return asPlainBytes(await packageDocx(patched, writeOptions(options)));
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
 * @param buffer - The source DOCX file.
 * @param patches - Array of patch operations to apply.
 * @param options - How to read the file (as `Document.read`: `password`, `securityPolicy`,
 *   `checkCrc32`) and write the result (as `Document.toBuffer`). `securityPolicy` applies to both.
 *   `password` is a read option here: the result of an encrypted file is not encrypted — read it with
 *   `Document.read` and write it with `Document.toBuffer(doc, { password })`, or `WordCrypto.encrypt` the result.
 * @returns New DOCX file as a Uint8Array.
 */
export async function patchDocument(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  patches: readonly PatchOperation[],
  options?: PatchOptions & ReadDocxOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "Package.patch";
  checkPatches(patches, where);
  checkPatchOptions(options, where);
  const doc = await readPackage(buffer, options, where);
  const patched = applyPatchesToDocument(doc, patches);
  return asPlainBytes(await packageDocx(patched, writeOptions(options)));
}

/** `Document.read` under the name of the `Package` function that reads. */
async function readPackage(
  buffer: unknown,
  options: ReadDocxOptions | undefined,
  where: string
): Promise<DocxDocument> {
  const bytes = checkBytes(buffer, "buffer", where);
  checkReadOptions(options, where);
  return readDocx(bytes, options);
}

/** The write side of a `Package` function's options, as `Document.toBuffer` takes them. */
function writeOptions(options: PatchOptions | undefined): DocxWriteOptions {
  return {
    compressionLevel: options?.compressionLevel,
    modTime: options?.modTime,
    securityPolicy: options?.securityPolicy
  };
}

// =============================================================================
// Template Engine - Convenience IO
// =============================================================================

/**
 * Read a DOCX buffer, fill template placeholders with data, and produce a new DOCX.
 *
 * This is a convenience wrapper combining readDocx + fillTemplate + packageDocx.
 *
 * @param buffer - The source DOCX template file as a Uint8Array.
 * @param data - Data object to fill into the template placeholders.
 * @param options - Template and compression options.
 * @returns New DOCX file as a Uint8Array.
 */
export async function fillTemplateFromBuffer(
  buffer: Uint8Array,
  data: Record<string, unknown>,
  options?: LegacyTemplateOptions & { readonly compressionLevel?: number }
): Promise<Uint8Array<ArrayBuffer>> {
  checkPackageOptions(options, "options", "Io.fillTemplateFromBuffer");
  const doc = await readDocx(buffer);
  const filled = fillTemplate(doc, data, options);
  return asPlainBytes(await packageDocx(filled, { compressionLevel: options?.compressionLevel }));
}

/**
 * Read a DOCX template, fill it with data and package the result — `Document.read`,
 * `Template.fill` and `Document.toBuffer` in one call, honouring every `fill` option
 * including `richPlaceholders`.
 *
 * @param buffer - The source DOCX template file.
 * @param data - Data to fill into the template placeholders: an object or a `Map`, as `Template.fill` takes it.
 * @param options - Template options, how to read the file (as `Document.read`: `password`,
 *   `securityPolicy`, `checkCrc32`) and how to write the result (as `Document.toBuffer`).
 *   `securityPolicy` applies to both. The result of an encrypted file is not
 *   encrypted: `WordCrypto.encrypt` it again.
 * @returns The filled DOCX file.
 */
export async function fillBuffer(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  data: Readonly<Record<string, unknown>> | ReadonlyMap<string, unknown>,
  options?: TemplateOptions & PatchOptions & ReadDocxOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "Package.fillTemplate";
  // Checked before the file is read, so a bad argument costs no parse.
  fillData(data, where);
  checkPatchOptions(options, where);
  const doc = await readPackage(buffer, options, where);
  return asPlainBytes(
    await packageDocx(fillDocument(doc, data, options, where), writeOptions(options))
  );
}

// =============================================================================
// Flat OPC Convenience
// =============================================================================

/** Options for {@link toFlatOpc}. */
export interface ToFlatOpcOptions {
  /**
   * Compression level of the intermediate package (0-9). Default: 0 — the
   * package is unzipped again immediately, so compressing it is wasted work.
   */
  readonly compressionLevel?: number;
}

/**
 * Package a DocxDocument model into Flat OPC XML format.
 *
 * This packages the document to a ZIP, then re-reads the entries to wrap
 * them as Flat OPC XML. While this involves a ZIP round-trip, it ensures
 * the Flat OPC output is byte-identical to what `Document.toBuffer` would produce
 * (same XML serialization, same relationships, same content types).
 *
 * @param doc - The document model.
 * @param options - Packaging options.
 * @returns The Flat OPC XML string.
 */
export async function toFlatOpc(doc: DocxDocument, options?: ToFlatOpcOptions): Promise<string> {
  checkHandle(doc, "document", "doc", "Document.toFlatOpc");
  checkSchema(toFlatOpcOptionsSchema.optional(), options, "Document.toFlatOpc");
  const zipBytes = await packageDocx(doc, { compressionLevel: options?.compressionLevel ?? 0 });

  // Statically imported: this file already imports `readDocx`, which imports `unzip` statically —
  // as do the ODT converter, the streaming reader and the incremental editor. A lone `await
  // import()` of a module four siblings pull eagerly saves nothing and only makes rolldown report
  // the pair as `INEFFECTIVE_DYNAMIC_IMPORT`.
  const reader = unzipZip(zipBytes);
  const entries = new Map<string, Uint8Array>();
  for await (const entry of reader.entries()) {
    const data = await entry.bytes();
    const path = entry.path.replace(/^\//, "").replace(/\\/g, "/");
    entries.set(path, data);
  }

  const { toFlatOpc: entriesToFlatOpc } = await import("@word/convert/flat-opc");
  return entriesToFlatOpc(entries);
}

/**
 * Package a DocxDocument model into Flat OPC XML format.
 *
 * @deprecated Use {@link toFlatOpc}, which takes an options object. Removed in 1.0.
 */
export async function toFlatOpcFromDoc(
  doc: DocxDocument,
  compressionLevel?: number
): Promise<string> {
  return toFlatOpc(doc, { compressionLevel });
}
