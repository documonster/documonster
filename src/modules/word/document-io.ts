/**
 * DOCX Module - Document IO
 *
 * IO operations: package/unpackage to bytes, base64, Flat OPC,
 * plus the public Patch and Template APIs that combine read/patch/write.
 */

import { bytesToBase64 } from "@word/core/internal-utils";
import type { PatchOperation } from "@word/patcher";
import { applyPatchesToDocument } from "@word/patcher";
import { readDocx } from "@word/reader/docx-reader";
import { fillTemplate } from "@word/template/template-engine";
import type { TemplateOptions } from "@word/template/template-engine";
import type { DocxDocument } from "@word/types";
import type { PackageDocxOptions } from "@word/writer/docx-packager";
import { packageDocx } from "@word/writer/docx-packager";

// Re-export patch types for backward compatibility
export type { PatchContent, PatchOperation } from "@word/patcher";

// =============================================================================
// Document IO (toBuffer / toBase64)
// =============================================================================

/** Package a DocxDocument model to DOCX bytes. */
export async function toBuffer(
  doc: DocxDocument,
  options?: PackageDocxOptions
): Promise<Uint8Array> {
  return packageDocx(doc, options);
}

/** Package a DocxDocument model to base64 string. */
export async function toBase64(doc: DocxDocument, options?: PackageDocxOptions): Promise<string> {
  const bytes = await toBuffer(doc, options);
  return bytesToBase64(bytes);
}

// =============================================================================
// Patcher / Template Fill API
// =============================================================================

/** Options for patchDocument. */
export interface PatchOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
}

/** A compiled template that can be reused for multiple patch operations. */
export interface CompiledTemplate {
  /** The parsed document model (internal use only). */
  readonly _doc: DocxDocument;
}

/**
 * Compile a DOCX template for reuse with multiple data sets.
 *
 * Parsing the ZIP and XML is expensive. If you need to patch the same template
 * multiple times with different data, compile it once and reuse.
 *
 * @param buffer - The source DOCX template file as a Uint8Array.
 * @returns A compiled template handle.
 */
export async function compileTemplate(buffer: Uint8Array): Promise<CompiledTemplate> {
  const doc = await readDocx(buffer);
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
 * @param options - Optional compression settings.
 * @returns New DOCX file as a Uint8Array.
 */
export async function patchTemplate(
  template: CompiledTemplate,
  patches: readonly PatchOperation[],
  options?: PatchOptions
): Promise<Uint8Array> {
  // Deep clone the document to avoid mutating the cached template
  const doc = structuredClone(template._doc);
  const patched = applyPatchesToDocument(doc, patches);
  return packageDocx(patched, { compressionLevel: options?.compressionLevel });
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
  const patched = applyPatchesToDocument(doc, patches);
  return packageDocx(patched, { compressionLevel: options?.compressionLevel });
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
  options?: TemplateOptions & { readonly compressionLevel?: number }
): Promise<Uint8Array> {
  const doc = await readDocx(buffer);
  const filled = fillTemplate(doc, data, options);
  return packageDocx(filled, { compressionLevel: options?.compressionLevel });
}

// =============================================================================
// Flat OPC Convenience
// =============================================================================

/**
 * Package a DocxDocument model into Flat OPC XML format.
 *
 * This packages the document to a ZIP, then re-reads the entries to wrap
 * them as Flat OPC XML. While this involves a ZIP round-trip, it ensures
 * the Flat OPC output is byte-identical to what `packageDocx` would produce
 * (same XML serialization, same relationships, same content types).
 *
 * @param doc - The document model.
 * @param compressionLevel - Optional compression level (0 = no compression, faster for Flat OPC).
 * @returns The Flat OPC XML string.
 */
export async function toFlatOpcFromDoc(
  doc: DocxDocument,
  compressionLevel?: number
): Promise<string> {
  // Use level 0 (store-only) since we're immediately decompressing
  const zipBytes = await packageDocx(doc, { compressionLevel: compressionLevel ?? 0 });

  const { unzip } = await import("@archive/read-archive");
  const reader = unzip(zipBytes);
  const entries = new Map<string, Uint8Array>();
  for await (const entry of reader.entries()) {
    const data = await entry.bytes();
    const path = entry.path.replace(/^\//, "").replace(/\\/g, "/");
    entries.set(path, data);
  }

  const { toFlatOpc } = await import("@word/convert/flat-opc");
  return toFlatOpc(entries);
}
