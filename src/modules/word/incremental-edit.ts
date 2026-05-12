/**
 * DOCX Incremental Edit API
 *
 * Supports efficient "open → modify → save" workflows by editing a DOCX
 * file at the ZIP entry level.
 *
 * Use cases:
 * - Replace just the document.xml in an existing template
 * - Update headers/footers without touching styles or images
 * - Patch specific parts (e.g. core properties metadata)
 *
 * **Limitations and caveats — read before using `replacePart`/`deletePart`:**
 *
 * This API is a thin wrapper around ZIP entry mutation. It deliberately
 * does not understand DOCX semantics:
 *
 * - `[Content_Types].xml` is **not** updated when you add a new part with
 *   a previously-unseen content type. If your replacement part introduces
 *   a new MIME type, callers must also patch `[Content_Types].xml` (or
 *   round-trip through `readDocx` + `packageDocx` instead).
 * - Relationship files (`*.rels`) are **not** rebuilt. Adding or deleting
 *   a part with relationships will leave the corresponding `_rels/*.rels`
 *   stale.
 * - `replaceBody`, `replaceHeader`, `replaceFooter` only rewrite their
 *   target XML part; they do **not** modify the corresponding `.rels` file
 *   or copy in new media binaries. As a result they reject content that
 *   would introduce new relationship IDs (images, hyperlinks, charts,
 *   altChunks, opaque drawings carrying rIds) — use `patchDocument` /
 *   `packageDocx` for those cases. Existing references that survive
 *   unchanged are fine because their `.rels` entries already exist.
 *
 * For DOCX-aware edits (changes that may add/remove relationships or
 * content types) use `patchDocument` or `compileTemplate`, or round-trip
 * through `readDocx` + `packageDocx`.
 *
 * Note: This API operates on raw ZIP entries; it does not parse the DOCX
 * model. For higher-level edits use `patchDocument` or `compileTemplate`.
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { XmlWriter } from "@xml/writer";

import { PartPath } from "./constants";
import { utf8Decoder, utf8Encoder } from "./core/internal-utils";
import { walkBlocks } from "./core/walker";
import { DocxUnsupportedFeatureError } from "./errors";
import type { BodyContent, DocxDocument, Paragraph, Table } from "./types";
import { renderBodyContent, renderDocument } from "./writer/document-writer";
import { renderHeader, renderFooter } from "./writer/header-footer-writer";
import { createRenderContext } from "./writer/render-context";

// =============================================================================
// Public API
// =============================================================================

/** A single incremental edit operation. */
export type IncrementalEdit =
  | {
      /**
       * Replace a specific part by raw bytes.
       *
       * Caveat: does not update `[Content_Types].xml` or any `.rels` file.
       * Adding a part with a new content type or new relationships will
       * produce a malformed package. See the file header for details.
       */
      readonly type: "replacePart";
      readonly path: string;
      readonly data: Uint8Array;
    }
  | {
      /** Replace a part with a string (UTF-8 encoded). Same caveats as `replacePart`. */
      readonly type: "replacePartText";
      readonly path: string;
      readonly text: string;
    }
  | {
      /**
       * Delete a specific part.
       *
       * Caveat: does not update `[Content_Types].xml` or any `.rels` file.
       * Deleting a part that other parts reference (via rels or other XML)
       * will leave dangling references — round-trip through readDocx +
       * packageDocx if you need a consistent package.
       */
      readonly type: "deletePart";
      readonly path: string;
    }
  | {
      /** Replace the body content of `word/document.xml`. */
      readonly type: "replaceBody";
      readonly body: readonly BodyContent[];
    }
  | {
      /** Replace a header part identified by path (e.g. `word/header1.xml`). */
      readonly type: "replaceHeader";
      readonly path: string;
      readonly children: readonly (Paragraph | Table)[];
    }
  | {
      /** Replace a footer part identified by path (e.g. `word/footer1.xml`). */
      readonly type: "replaceFooter";
      readonly path: string;
      readonly children: readonly (Paragraph | Table)[];
    };

/** Options for incremental editing. */
export interface IncrementalEditOptions {
  /** Compression level for replaced entries (0-9). Default: 6. */
  readonly compressionLevel?: number;
}

/**
 * Apply incremental edits to an existing DOCX file.
 *
 * Reads all parts from the original ZIP, applies the edits, and writes a
 * new ZIP. Unchanged parts are passed through (decompressed and recompressed,
 * but their content is not parsed).
 *
 * @param buffer - The original DOCX file as a Uint8Array.
 * @param edits - Array of incremental edits to apply.
 * @param options - Optional settings.
 * @returns A new DOCX file as a Uint8Array.
 *
 * @example
 * ```ts
 * const original = await fs.readFile("template.docx");
 * const edited = await editDocxIncremental(original, [
 *   { type: "replaceBody", body: [...newParagraphs] }
 * ]);
 * ```
 */
export async function editDocxIncremental(
  buffer: Uint8Array,
  edits: readonly IncrementalEdit[],
  options?: IncrementalEditOptions
): Promise<Uint8Array> {
  // Read all entries from the original ZIP. Some third-party ZIP writers emit
  // entry paths with a leading "/" — strip it so lookups by canonical path
  // (e.g. "word/document.xml") match consistently with the rest of the
  // codebase (see readDocxPart below).
  const reader = unzip(buffer);
  const parts = new Map<string, Uint8Array>();
  for await (const entry of reader.entries()) {
    const normalizedPath = entry.path.replace(/^\/+/, "");
    parts.set(normalizedPath, await entry.bytes());
  }

  // Apply edits
  for (const edit of edits) {
    switch (edit.type) {
      case "replacePart":
        parts.set(edit.path.replace(/^\/+/, ""), edit.data);
        break;
      case "replacePartText":
        parts.set(edit.path.replace(/^\/+/, ""), utf8Encoder.encode(edit.text));
        break;
      case "deletePart":
        parts.delete(edit.path.replace(/^\/+/, ""));
        break;
      case "replaceBody": {
        // Preserve the original document's section properties, background,
        // and document-level wrappers — only swap the children inside w:body.
        // Reject new external references for the same reason as
        // replaceHeader/Footer: this editor doesn't rewrite document.xml.rels.
        assertNoExternalReferences(edit.body, "body");
        const original = parts.get(PartPath.Document);
        parts.set(
          PartPath.Document,
          original
            ? replaceBodyChildrenInDocumentXml(original, edit.body)
            : renderBodyOnly(edit.body)
        );
        break;
      }
      case "replaceHeader":
        parts.set(edit.path.replace(/^\/+/, ""), renderHeaderFooter("header", edit.children));
        break;
      case "replaceFooter":
        parts.set(edit.path.replace(/^\/+/, ""), renderHeaderFooter("footer", edit.children));
        break;
    }
  }

  // Write back to a new ZIP
  const archive = zip({ level: options?.compressionLevel ?? 6 });
  for (const [path, data] of parts) {
    archive.add(path, data);
  }
  return archive.bytes();
}

/**
 * List all part paths in a DOCX file without parsing the model.
 *
 * @param buffer - The DOCX file as a Uint8Array.
 * @returns Array of part paths.
 */
export async function listDocxParts(buffer: Uint8Array): Promise<string[]> {
  const reader = unzip(buffer);
  const paths: string[] = [];
  for await (const entry of reader.entries()) {
    paths.push(entry.path);
  }
  return paths;
}

/**
 * Read a single part's raw bytes from a DOCX file.
 *
 * @param buffer - The DOCX file as a Uint8Array.
 * @param path - The part path (e.g. `"word/document.xml"`).
 * @returns The part's raw bytes, or undefined if the part doesn't exist.
 */
export async function readDocxPart(
  buffer: Uint8Array,
  path: string
): Promise<Uint8Array | undefined> {
  const reader = unzip(buffer);
  for await (const entry of reader.entries()) {
    const normalizedPath = entry.path.replace(/^\//, "");
    if (normalizedPath === path) {
      return entry.bytes();
    }
  }
  return undefined;
}

// =============================================================================
// Internal helpers
// =============================================================================

function renderBodyOnly(body: readonly BodyContent[]): Uint8Array {
  const minimalDoc: DocxDocument = { body } as DocxDocument;
  const ctx = createRenderContext();

  const xml = new XmlWriter();
  renderDocument(xml, minimalDoc, ctx);
  return utf8Encoder.encode(xml.toString());
}

/**
 * Replace the children inside `<w:body>...</w:body>` of an existing
 * document.xml, preserving the original section properties (the trailing
 * `w:sectPr` directly under w:body), the document background, and any other
 * document-level structure (namespace declarations, custom attributes, etc.).
 *
 * If the document doesn't contain a parseable `<w:body>` block we fall back
 * to a minimal render so callers still get a valid DOCX (this matches the
 * legacy behavior of `replaceBody`).
 */
function replaceBodyChildrenInDocumentXml(
  original: Uint8Array,
  body: readonly BodyContent[]
): Uint8Array {
  const text = utf8Decoder.decode(original);

  // Locate <w:body...> opening tag and the matching </w:body>. We intentionally
  // search from the end for the closing tag to tolerate stray "</w:body>"
  // mentions inside CDATA or comments — neither of which Word emits, but we
  // stay defensive.
  const bodyOpen = /<w:body(?:\s[^>]*)?>/.exec(text);
  const bodyCloseIdx = text.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyCloseIdx < 0 || bodyOpen.index >= bodyCloseIdx) {
    return renderBodyOnly(body);
  }

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const innerBody = text.slice(bodyStart, bodyCloseIdx);

  // Preserve the trailing <w:sectPr ...>...</w:sectPr> if present. Section
  // properties at this level apply to the whole document and would otherwise
  // be lost.
  let trailingSectPr = "";
  const sectPrMatch =
    /<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>\s*$|<w:sectPr(?:\s[^>]*)?\/>\s*$/.exec(innerBody);
  if (sectPrMatch) {
    trailingSectPr = sectPrMatch[0].trimEnd();
  }

  // Render new body children (no document/body wrappers, no sectPr — we add
  // back the preserved one ourselves).
  const ctx = createRenderContext();
  const xml = new XmlWriter();
  for (const block of body) {
    renderBodyContent(xml, block, ctx);
  }
  const newInner = xml.toString() + trailingSectPr;

  const updated = text.slice(0, bodyStart) + newInner + text.slice(bodyCloseIdx);
  return utf8Encoder.encode(updated);
}

function renderHeaderFooter(
  kind: "header" | "footer",
  children: readonly (Paragraph | Table)[]
): Uint8Array {
  // Reject content that requires fresh part-level relationships (images,
  // external hyperlinks, charts, alt-chunks, embedded fonts, etc.). The
  // incremental editor only rewrites individual XML parts; it does not
  // rebuild the per-part `*.rels` file or copy in new media binaries, so
  // any new `r:embed` / `r:id` reference would dangle. Callers wanting to
  // add such content should use the full `patchDocument` / `packageDocx`
  // flow instead.
  assertNoExternalReferences(children, kind);

  const xml = new XmlWriter();
  if (kind === "header") {
    renderHeader(xml, { children });
  } else {
    renderFooter(xml, { children });
  }
  return utf8Encoder.encode(xml.toString());
}

/**
 * Walk a header/footer / body fragment and throw if it contains any element
 * that needs a per-part relationship registration the incremental editor
 * cannot perform.
 *
 * The list mirrors what {@link renderBodyContent} writes as `r:embed` /
 * `r:id` / `r:link` attributes:
 *   - inline & floating images (`r:embed` / `r:link`)
 *   - external hyperlinks (`r:id` on `w:hyperlink`)
 *   - charts and chart extensions (`r:id` on `w:chart` / `cx:chart`)
 *   - alt-chunks (`r:id` on `w:altChunk`)
 *   - opaque drawings carrying preserved relationships
 */
function assertNoExternalReferences(blocks: readonly BodyContent[], context: string): void {
  let offending: string | undefined;
  walkBlocks(blocks, {
    enterRun(run) {
      for (const c of run.content) {
        if (c.type === "image") {
          offending ??= "inline image";
          return "stop";
        }
      }
      return "continue";
    },
    enterHyperlink(h) {
      if (h.url || h.rId) {
        offending ??= "hyperlink";
        return "stop";
      }
      return "continue";
    },
    visitFloatingImage() {
      offending ??= "floating image";
    },
    visitChart() {
      offending ??= "chart";
    },
    visitChartEx() {
      offending ??= "chart (extended)";
    },
    visitAltChunk() {
      offending ??= "altChunk";
    },
    visitOpaqueDrawing(drawing) {
      if (drawing.referencedRIds.length > 0) {
        offending ??= "opaque drawing with relationships";
      }
    }
  });

  if (offending) {
    throw new DocxUnsupportedFeatureError(
      `incremental ${context} replacement does not support ${offending}: ` +
        "relationship parts (.rels) are not rewritten. Use patchDocument() or " +
        "rebuild the document with packageDocx() instead."
    );
  }
}
